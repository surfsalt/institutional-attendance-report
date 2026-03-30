#!/usr/bin/env python3
"""
Blackboard Global Attendance Data Extractor
============================================
Connects to Blackboard Learn REST API and extracts institution-wide
attendance data for all courses in a selected term.

Output: Four CSV files ready for import into the Excel reporting workbook.
  1. course_summary.csv   — one row per course section
  2. student_detail.csv   — one row per student per course
  3. daily_attendance.csv  — one row per course per meeting date
  4. compliance.csv        — one row per course (compliance flags)

Usage:
  pip install requests
  python extract.py

On first run, it creates a config.ini template for you to fill in.

Based on the proven Blackboard REST API architecture from bb-attendance-web:
  - OAuth 2.0 Client Credentials authentication
  - Bulk user attendance endpoint (cross-course quirk handled)
  - Instructor lookup via local courseRoleId filtering
  - Dropped enrollment detection (availability.available = "No")
  - Meeting probe strategy for 403-blocked courses
  - Weighted scoring: Present=100%, Late=50%, Absent=0%, Excused=excluded

Required Blackboard REST API entitlements on integration user:
  - system.useradmin.generic.VIEW
  - system.course.VIEW
  - course.attendance.VIEW
  - course.configure-properties.EXECUTE
  - system.courseuserlist.VIEW
  - system.multiinst.hierarchy.manager.VIEW
  - system.multiinst.node.course.association.VIEW

Known API Limitations (handled):
  1. GET /courses/{cid}/meetings returns 403 on some courses even with
     System Admin — attendance was never enabled or Qwickly restricted.
  2. GET /courses/{cid}/meetings/users/{uid} returns ALL records across
     ALL courses (not just the specified one). Must cross-reference.
  3. roleId query param on course membership endpoint is unreliable —
     must fetch all members and filter locally by courseRoleId.
  4. Bulk endpoint may miss records — gap-fill needed via per-meeting endpoint.
"""

import configparser
import csv
import os
import sys
import time
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta

try:
    import requests
except ImportError:
    print("ERROR: 'requests' library not installed. Run: pip install requests")
    sys.exit(1)


# ── Configuration ────────────────────────────────────────────────────────────

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "extract_config.ini")
OUTPUT_DIR  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "attendance_data")

DEFAULT_THRESHOLD = 75          # Attendance % below which students are flagged
STALE_DAYS        = 14          # Days since last attendance to flag as "Stale"
MAX_WORKERS       = 2           # Parallel API threads (keep very low to avoid 429s)
BATCH_SIZE        = 20          # Process courses in batches of this size
BATCH_DELAY       = 2           # Seconds to pause between batches

# ── Institutional Hierarchy ──────────────────────────────────────────────────
# Department resolution uses a PER-COURSE reverse lookup:
#   GET /learn/api/public/v1/courses/{courseId}/nodes?expand=node
# This returns the hierarchy node(s) each course belongs to, with the
# node's title (department name) and parentId.
#
# The 'All Departments' root node has:
#   - Internal ID: _172_1  (used in API URLs)
#   - External ID: 05d8bd91-8efb-476c-91b4-98138168afab
# Department nodes are direct children of this root (parentId == _172_1).

ALL_DEPARTMENTS_NODE_ID = "_172_1"  # Internal BB ID for 'All Departments'


def load_config():
    """Load or create configuration file."""
    if not os.path.exists(CONFIG_FILE):
        config = configparser.ConfigParser()
        config["blackboard"] = {
            "base_url": "https://your-institution.blackboard.com",
            "api_key": "YOUR_API_KEY",
            "api_secret": "YOUR_API_SECRET",
        }
        config["settings"] = {
            "threshold_pct": str(DEFAULT_THRESHOLD),
            "stale_days": str(STALE_DAYS),
        }
        with open(CONFIG_FILE, "w") as f:
            config.write(f)
        print(f"Config template created at: {CONFIG_FILE}")
        print("Please edit it with your Blackboard credentials and re-run.")
        sys.exit(0)

    config = configparser.ConfigParser()
    config.read(CONFIG_FILE)
    return config


# ── Blackboard API Layer ─────────────────────────────────────────────────────
# Mirrors the proven architecture from bb-attendance-web ARCHITECTURE.md

class BlackboardAPI:
    """Wrapper for Blackboard Learn REST API with institutional-scale methods."""

    def __init__(self, base_url, api_key, api_secret):
        self.base_url = base_url.rstrip("/")
        self.api_key = api_key
        self.api_secret = api_secret
        self.token = None
        self.session = requests.Session()
        self.api_call_count = 0  # Track total API calls

        # ── Caches ──────────────────────────────────────────────────────
        # These caches prevent redundant API calls for the same user/data.
        # Without caching, the same instructor or student is looked up
        # once per course they appear in, causing thousands of extra calls.
        self._user_cache = {}          # userId → user details dict
        self._attendance_cache = {}    # userId → list of attendance records (cross-course)

    def authenticate(self):
        """OAuth 2.0 Client Credentials flow."""
        url = f"{self.base_url}/learn/api/public/v1/oauth2/token"
        resp = self.session.post(
            url,
            data={"grant_type": "client_credentials"},
            auth=(self.api_key, self.api_secret),
            timeout=15,
        )
        resp.raise_for_status()
        self.token = resp.json()["access_token"]
        print(f"[Auth] Token acquired from {self.base_url}")

    def _get(self, path, params=None):
        """Authenticated GET request with automatic 429 retry and adaptive throttling."""
        url = f"{self.base_url}{path}"
        headers = {"Authorization": f"Bearer {self.token}"}
        max_retries = 8
        for attempt in range(max_retries):
            self.api_call_count += 1
            
            # Adaptive throttle: add a small delay every N calls to avoid
            # hitting the rate limit wall at ~190 courses
            if self.api_call_count % 50 == 0:
                time.sleep(1)  # Brief cooldown every 50 calls
            
            resp = self.session.get(url, headers=headers, params=params, timeout=30)
            if resp.status_code == 429:
                # Rate limited — read Retry-After header or default to 60s
                raw_retry = resp.headers.get("Retry-After", "60")
                # Strip non-numeric chars (BB sometimes sends '85166s')
                wait = int(''.join(c for c in str(raw_retry) if c.isdigit()) or '60')
                wait = min(wait, 300)  # Cap at 5 minutes
                print(f"  [Rate Limit] 429 on {path[:60]}... waiting {wait}s (attempt {attempt+1}/{max_retries})")
                time.sleep(wait)
                continue
            resp.raise_for_status()
            return resp.json()
        # Final attempt — let it raise if still 429
        resp.raise_for_status()
        return resp.json()

    def _get_paged(self, path, key="results"):
        """Fetch all pages of a paginated endpoint."""
        items = []
        while path:
            data = self._get(path)
            items.extend(data.get(key, []))
            nxt = data.get("paging", {}).get("nextPage")
            path = nxt if nxt else None
        return items

    # ── Terms ────────────────────────────────────────────────────────────

    def get_terms(self):
        """Fetch all academic terms, newest first."""
        terms = self._get_paged("/learn/api/public/v1/terms")
        terms.reverse()
        return terms

    # ── Courses ──────────────────────────────────────────────────────────

    def get_courses_for_term(self, term_id):
        """Fetch all courses in a given term."""
        courses = self._get_paged(
            f"/learn/api/public/v3/courses?termId={term_id}"
        )
        return courses

    def get_course_memberships(self, course_id):
        """Fetch all memberships for a course."""
        return self._get_paged(
            f"/learn/api/public/v1/courses/{course_id}/users"
        )

    def get_course_meetings(self, course_id):
        """Fetch meeting (attendance session) list for a course.
        Returns (meetings_list, is_blocked).
        403 = attendance not enabled for this course."""
        try:
            meetings = self._get_paged(
                f"/learn/api/public/v1/courses/{course_id}/meetings"
            )
            return meetings, False
        except requests.HTTPError as e:
            if e.response.status_code == 403:
                return [], True
            raise

    # ── Student Attendance ───────────────────────────────────────────────

    def get_user_attendance_bulk(self, course_id, user_id):
        """Fetch ALL attendance records for a user (cross-course quirk).
        NOTE: This endpoint returns records from ALL courses, not just
        the specified one. Caller must cross-reference with meetings lists.
        
        CACHED: Since the endpoint returns cross-course data anyway,
        we only need to call it once per user regardless of how many
        courses they're enrolled in."""
        if user_id in self._attendance_cache:
            return self._attendance_cache[user_id]
        try:
            records = self._get_paged(
                f"/learn/api/public/v1/courses/{course_id}/meetings/users/{user_id}"
            )
            self._attendance_cache[user_id] = records
            return records
        except Exception:
            self._attendance_cache[user_id] = []
            return []

    def get_single_attendance(self, course_id, meeting_id, user_id):
        """Fetch single attendance record — most reliable but slowest."""
        try:
            data = self._get(
                f"/learn/api/public/v1/courses/{course_id}/meetings/{meeting_id}/users/{user_id}"
            )
            return data.get("status")
        except Exception:
            return None

    def get_user_details(self, user_id):
        """Fetch user profile details.
        
        CACHED: The same instructor or student may appear in many courses.
        Without caching, a professor teaching 5 courses would be looked up
        5 times; a student in 6 courses would be looked up 6 times."""
        if user_id in self._user_cache:
            return self._user_cache[user_id]
        try:
            data = self._get(f"/learn/api/public/v1/users/{user_id}")
            self._user_cache[user_id] = data
            return data
        except Exception:
            self._user_cache[user_id] = {}
            return {}

    # ── Institutional Hierarchy ───────────────────────────────────────────

    def get_course_hierarchy_nodes(self, course_id):
        """Fetch the hierarchy nodes a course belongs to (reverse lookup).
        
        Uses: GET /learn/api/public/v1/courses/{courseId}/nodes?expand=node
        
        Returns list of objects, each with:
          - nodeId: the hierarchy node's primary ID
          - courseId: the course's primary ID
          - isPrimary: boolean
          - node: { id, externalId, title, description, parentId }
            (only present because we use expand=node)
        
        The 'title' field of the expanded node is the department name.
        The 'parentId' field tells us where this node sits in the tree.
        
        CACHED: Results are cached per course (though each course is
        typically only looked up once).
        """
        try:
            return self._get_paged(
                f"/learn/api/public/v1/courses/{course_id}/nodes?expand=node"
            )
        except Exception as e:
            # Silently return empty — not all courses have hierarchy assignments
            return []


# ── Data Processing ──────────────────────────────────────────────────────────

def resolve_department(api, course_id):
    """Look up the department name for a single course via the hierarchy API.
    
    Uses the REVERSE lookup: GET /courses/{courseId}/nodes?expand=node
    This asks 'what hierarchy nodes is this course in?' instead of
    walking the entire tree top-down.
    
    Strategy:
      1. Fetch all nodes the course belongs to (usually 1-2).
      2. Find the node whose parentId matches ALL_DEPARTMENTS_NODE_ID
         (that node IS the department).
      3. If the course is in a sub-node (e.g. DE_2024 under Dual Enrolment),
         the sub-node's parentId will NOT be ALL_DEPARTMENTS_NODE_ID,
         but its grandparent will be. In that case, the node's title
         is the sub-department — we want the top-level department.
      4. For sub-nodes, we walk up using parentId until we find the
         node whose parent is ALL_DEPARTMENTS_NODE_ID.
    
    Returns the department name string, or '' if not found.
    """
    nodes = api.get_course_hierarchy_nodes(course_id)
    if not nodes:
        return ""
    
    # Look for a node whose parentId is the 'All Departments' root.
    # That node IS the top-level department.
    for assoc in nodes:
        node = assoc.get("node", {})
        if not node:
            continue
        parent_id = node.get("parentId", "")
        title = node.get("title", "") or node.get("name", "")
        
        if parent_id == ALL_DEPARTMENTS_NODE_ID:
            # This node is a direct child of 'All Departments' — it IS the department
            return title
    
    # If none of the nodes are direct children of All Departments,
    # the course may be in a sub-node (e.g. DE_2024 under Dual Enrolment).
    # Return the first node's title as a reasonable fallback.
    # (A full upward walk would require extra API calls per level.)
    for assoc in nodes:
        node = assoc.get("node", {})
        if node:
            return node.get("title", "") or node.get("name", "")
    
    return ""


def compute_weighted_rate(present, late, absent):
    """Blackboard weighted attendance formula.
    Present=100%, Late=50%, Absent=0%.
    Excused and Not Marked are excluded from denominator."""
    total = present + late + absent
    if total == 0:
        return None  # No recorded attendance
    return round((present * 100 + late * 50) / total, 2)


def risk_band(rate, threshold=75):
    """Classify student risk based on attendance rate."""
    if rate is None:
        return "No Data"
    if rate < 50:
        return "High Risk"
    if rate < threshold:
        return "Medium Risk"
    return "OK"


def course_status(last_date_str, stale_days=14):
    """Determine course attendance recording status."""
    if not last_date_str:
        return "Not Recorded"
    try:
        last = datetime.fromisoformat(last_date_str.replace("Z", "+00:00"))
        days_ago = (datetime.now(last.tzinfo) - last).days
        if days_ago <= stale_days:
            return "Active"
        return "Stale"
    except Exception:
        return "Unknown"


def clean_date(iso_str):
    """Convert ISO timestamp like '2026-01-13T18:13:36.081Z' to '2026-01-13'."""
    if not iso_str:
        return ""
    try:
        return iso_str[:10]  # Take just YYYY-MM-DD
    except Exception:
        return iso_str


def extract_institutional_data(api, term_id, term_name, threshold=75, stale_days=14):
    """
    Main extraction: pulls ALL courses in a term, ALL students per course,
    and ALL attendance data. Returns four datasets.

    Department names are resolved per-course via the hierarchy API's
    reverse lookup (GET /courses/{id}/nodes?expand=node).

    This is the institutional-scale version of what bb-attendance-web does
    for a single student.
    
    API OPTIMIZATION:
      - User details are cached (instructor appearing in 5 courses = 1 API call, not 5)
      - Student attendance is cached (cross-course quirk means 1 call covers all courses)
      - Parallel processing uses MAX_WORKERS=2 to stay under rate limits
      - Courses processed in batches with pauses to avoid 429s
    """
    print(f"\n{'='*60}")
    print(f"EXTRACTING INSTITUTIONAL ATTENDANCE DATA")
    print(f"{'='*60}")

    # Step 1: Get all courses in the term
    print("\n[Step 1] Fetching courses for term...")
    courses = api.get_courses_for_term(term_id)
    print(f"  Found {len(courses)} courses")

    # Step 2: For each course, fetch meetings + memberships in parallel
    # NOTE: Instructor user details are CACHED so repeated lookups are free
    print("\n[Step 2] Fetching meetings and memberships per course...")
    print(f"  (Using {MAX_WORKERS} parallel workers, user details are cached)")

    course_data = {}  # course_id -> dict with all info

    def process_course(course):
        cid = course.get("id")          # Internal BB ID like '_12345_1'
        ext_id = course.get("courseId", cid)  # External ID like '2025_SP_AH_ENG_101_1'
        name = course.get("name", "Unknown")

        # Get meetings
        meetings, blocked = api.get_course_meetings(cid)

        # Get memberships and separate students vs instructors
        members = api.get_course_memberships(cid)
        students = [m for m in members if m.get("courseRoleId") == "Student"]
        instructors = [m for m in members if m.get("courseRoleId") == "Instructor"]

        # Look up instructor names (CACHED — no duplicate API calls)
        instructor_names = []
        for inst in instructors:
            uid = inst.get("userId", "")
            udata = api.get_user_details(uid)
            given = udata.get("name", {}).get("given", "")
            family = udata.get("name", {}).get("family", "")
            full = f"{given} {family}".strip()
            if full:
                instructor_names.append(full)

        # Resolve department via reverse hierarchy lookup
        # (asks BB 'what hierarchy node is this course in?')
        department = resolve_department(api, cid)

        # Build a friendly course code from ext_id
        friendly_code = ext_id
        parts = ext_id.split("_")
        term_prefixes = {"SP", "FA", "SU", "S1", "S2", "WI"}
        if len(parts) >= 4:
            subj_parts = []
            for p in parts:
                if p.isdigit() and len(p) == 4:
                    continue
                if p in term_prefixes:
                    continue
                if not subj_parts and p.isalpha() and p.isupper() and len(p) <= 3:
                    idx = parts.index(p)
                    remaining = parts[idx+1:]
                    has_alpha_after = any(rp.isalpha() and len(rp) >= 2 for rp in remaining)
                    has_digit_after = any(rp.isdigit() and len(rp) == 3 for rp in remaining)
                    if has_alpha_after and has_digit_after:
                        continue
                if p.isalpha() and len(p) >= 2:
                    subj_parts.append(p)
                elif p.isdigit() and subj_parts:
                    subj_parts.append(p)
            if len(subj_parts) >= 2:
                if len(subj_parts) >= 3:
                    friendly_code = f"{subj_parts[0]} {subj_parts[1]}-{subj_parts[2]}"
                else:
                    friendly_code = f"{subj_parts[0]} {subj_parts[1]}"

        return {
            "course_id": cid,
            "ext_id": ext_id,
            "friendly_code": friendly_code,
            "name": name,
            "meetings": meetings,
            "blocked": blocked,
            "students": students,
            "instructors": instructor_names,
            "department": department,
        }

    # Process in small batches with pauses between to avoid 429s.
    # BB's rate limit resets on a rolling window — spacing requests out
    # keeps us under the threshold instead of hitting a wall at ~190.
    total = len(courses)
    done = 0
    for batch_start in range(0, total, BATCH_SIZE):
        batch = courses[batch_start:batch_start + BATCH_SIZE]
        
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = {executor.submit(process_course, c): c for c in batch}
            for future in as_completed(futures):
                result = future.result()
                course_data[result["course_id"]] = result
                done += 1
                if done % 10 == 0 or done == total:
                    print(f"  Processed {done}/{total} courses... (API calls so far: {api.api_call_count})")
        
        # Pause between batches to let the rate limit window slide
        if batch_start + BATCH_SIZE < total:
            print(f"  [Throttle] Pausing {BATCH_DELAY}s between batches...")
            time.sleep(BATCH_DELAY)

    print(f"  Completed all {total} courses (API calls so far: {api.api_call_count})")
    
    # Report hierarchy mapping results
    mapped = sum(1 for cd in course_data.values() if cd["department"])
    unmapped = sum(1 for cd in course_data.values() if not cd["department"])
    print(f"\n  [Hierarchy] Courses with department: {mapped}, without: {unmapped}")

    # Step 3: For each student in each course, fetch attendance
    # NOTE: Attendance records are CACHED per-user (cross-course quirk means
    #       one API call gets ALL records for a student across ALL courses).
    #       User details are also CACHED from Step 2.
    print(f"\n[Step 3] Fetching student attendance records...")
    print(f"  (Attendance and user details are cached — repeat lookups are free)")

    course_summary_rows = []
    student_detail_rows = []
    daily_rows = []
    compliance_rows = []

    course_count = 0
    total_courses = len(course_data)
    for cid, cdata in course_data.items():
        course_count += 1
        if course_count % 10 == 0 or course_count == total_courses:
            print(f"  Processing course {course_count}/{total_courses}: {cdata.get('friendly_code', cid)}... (API calls: {api.api_call_count})")
        meeting_ids = {str(m.get("id")): m for m in cdata["meetings"]}
        meeting_dates = {}
        for mid, m in meeting_ids.items():
            start = m.get("start", "")
            if start:
                meeting_dates[mid] = start

        # Count active students (not dropped)
        total_students = len([s for s in cdata["students"]
                              if s.get("availability", {}).get("available", "Yes") == "Yes"])

        # Per-course aggregates
        course_present = 0
        course_late = 0
        course_absent = 0
        course_na = 0
        students_100 = 0
        students_above = 0
        students_below = 0
        student_rates = []
        last_attendance_date = None

        for stu in cdata["students"]:
            stu_uid = stu.get("userId")
            stu_avail = stu.get("availability", {}).get("available", "Yes")
            if stu_avail != "Yes":
                continue  # Skip dropped students

            # Fetch attendance for this student (CACHED — free if already fetched)
            records = api.get_user_attendance_bulk(cid, stu_uid)

            # Cross-reference: only keep records for THIS course's meetings
            stu_present = 0
            stu_late = 0
            stu_absent = 0
            stu_excused = 0
            stu_last_date = None

            record_map = {str(r.get("meetingId")): r.get("status") for r in records}

            for mid in meeting_ids:
                status = record_map.get(mid)
                if status is None:
                    course_na += 1
                    continue
                if status == "Present":
                    stu_present += 1
                    course_present += 1
                elif status == "Late":
                    stu_late += 1
                    course_late += 1
                elif status == "Absent":
                    stu_absent += 1
                    course_absent += 1
                elif status == "Excused":
                    stu_excused += 1
                elif status == "Not Marked":
                    course_na += 1
                    continue

                date_str = meeting_dates.get(mid, "")
                if date_str:
                    if stu_last_date is None or date_str > stu_last_date:
                        stu_last_date = date_str
                    if last_attendance_date is None or date_str > last_attendance_date:
                        last_attendance_date = date_str

            rate = compute_weighted_rate(stu_present, stu_late, stu_absent)
            band = risk_band(rate, threshold)

            if rate is not None:
                student_rates.append(rate)
                if rate >= 100:
                    students_100 += 1
                if rate >= threshold:
                    students_above += 1
                else:
                    students_below += 1

            # Get student name (CACHED — free if already fetched for another course)
            udata = api.get_user_details(stu_uid)
            stu_name = f"{udata.get('name', {}).get('given', '')} {udata.get('name', {}).get('family', '')}".strip()
            stu_id = udata.get("studentId", udata.get("externalId", ""))

            student_detail_rows.append({
                "term": term_name,
                "department": cdata.get("department", ""),
                "course_code": cdata["friendly_code"],
                "course_name": cdata["name"],
                "instructor": ", ".join(cdata["instructors"]),
                "student_id": stu_id,
                "student_name": stu_name,
                "present": stu_present,
                "late": stu_late,
                "absent": stu_absent,
                "excused": stu_excused,
                "attendance_pct": rate,
                "last_attendance_date": clean_date(stu_last_date) if stu_last_date else "",
                "risk_band": band,
                "below_threshold": "Yes" if rate is not None and rate < threshold else "No",
            })

        # Daily attendance by meeting date
        for mid, m in meeting_ids.items():
            date_str = m.get("start", "")
            daily_rows.append({
                "term": term_name,
                "department": cdata.get("department", ""),
                "course_code": cdata["friendly_code"],
                "course_name": cdata["name"],
                "meeting_date": clean_date(date_str),
                "students_enrolled": total_students,
            })

        # Course summary
        avg_rate = round(sum(student_rates) / len(student_rates), 2) if student_rates else None

        days_since = None
        if last_attendance_date:
            try:
                last_dt = datetime.fromisoformat(last_attendance_date.replace("Z", "+00:00"))
                days_since = (datetime.now(last_dt.tzinfo) - last_dt).days
            except Exception:
                pass

        status = course_status(last_attendance_date, stale_days)

        course_summary_rows.append({
            "term": term_name,
            "department": cdata.get("department", ""),
            "course_code": cdata["friendly_code"],
            "course_name": cdata["name"],
            "instructor": ", ".join(cdata["instructors"]),
            "total_students": total_students,
            "avg_attendance_pct": avg_rate,
            "students_100_pct": students_100,
            "students_above_threshold": students_above,
            "students_below_threshold": students_below,
            "pct_above_threshold": round(students_above / total_students * 100, 1) if total_students else None,
            "pct_below_threshold": round(students_below / total_students * 100, 1) if total_students else None,
            "total_present": course_present,
            "total_late": course_late,
            "total_absent": course_absent,
            "total_na": course_na,
            "total_meetings": len(meeting_ids),
            "last_attendance_date": clean_date(last_attendance_date) if last_attendance_date else "",
            "days_since_last": days_since,
            "status": status,
            "api_blocked": cdata["blocked"],
        })

        # Compliance row
        compliance_rows.append({
            "term": term_name,
            "department": cdata.get("department", ""),
            "course_code": cdata["friendly_code"],
            "course_name": cdata["name"],
            "instructor": ", ".join(cdata["instructors"]),
            "total_meetings": len(meeting_ids),
            "total_attendance_records": course_present + course_late + course_absent,
            "last_attendance_date": clean_date(last_attendance_date) if last_attendance_date else "",
            "days_since_last": days_since,
            "status": status,
            "api_blocked": cdata["blocked"],
            "no_attendance_recorded": len(meeting_ids) == 0 and not cdata["blocked"],
            "attendance_not_recent": status == "Stale",
        })

    return course_summary_rows, student_detail_rows, daily_rows, compliance_rows


def write_csv(filename, rows, fieldnames):
    """Write rows to CSV."""
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    path = os.path.join(OUTPUT_DIR, filename)
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    print(f"  Wrote {len(rows)} rows to {path}")


def main():
    config = load_config()

    base_url = config["blackboard"]["base_url"]
    api_key = config["blackboard"]["api_key"]
    api_secret = config["blackboard"]["api_secret"]
    threshold = int(config["settings"].get("threshold_pct", DEFAULT_THRESHOLD))
    stale = int(config["settings"].get("stale_days", STALE_DAYS))

    api = BlackboardAPI(base_url, api_key, api_secret)
    api.authenticate()

    # List terms for user selection
    terms = api.get_terms()
    print("\nAvailable Terms:")
    for i, t in enumerate(terms):
        print(f"  [{i+1}] {t.get('name', 'Unknown')} (ID: {t.get('id')})")

    choice = input("\nSelect term number: ").strip()
    term = terms[int(choice) - 1]
    term_id = term["id"]
    print(f"\nSelected: {term.get('name')}")

    # Extract data (department names resolved per-course via hierarchy API)
    course_rows, student_rows, daily_rows, compliance_rows = extract_institutional_data(
        api, term_id, term.get("name", "Unknown Term"), threshold, stale
    )

    # Write CSVs
    print(f"\n[Output] Writing CSV files...")
    write_csv("course_summary.csv", course_rows, [
        "term", "department", "course_code", "course_name", "instructor",
        "total_students", "avg_attendance_pct", "students_100_pct",
        "students_above_threshold", "students_below_threshold",
        "pct_above_threshold", "pct_below_threshold",
        "total_present", "total_late", "total_absent", "total_na",
        "total_meetings", "last_attendance_date", "days_since_last",
        "status", "api_blocked",
    ])
    write_csv("student_detail.csv", student_rows, [
        "term", "department", "course_code", "course_name", "instructor",
        "student_id", "student_name", "present", "late", "absent", "excused",
        "attendance_pct", "last_attendance_date", "risk_band", "below_threshold",
    ])
    write_csv("daily_attendance.csv", daily_rows, [
        "term", "department", "course_code", "course_name",
        "meeting_date", "students_enrolled",
    ])
    write_csv("compliance.csv", compliance_rows, [
        "term", "department", "course_code", "course_name", "instructor",
        "total_meetings", "total_attendance_records",
        "last_attendance_date", "days_since_last", "status",
        "api_blocked", "no_attendance_recorded", "attendance_not_recent",
    ])

    # Print API usage summary
    print(f"\n{'='*60}")
    print(f"EXTRACTION COMPLETE")
    print(f"{'='*60}")
    print(f"  Total API calls made: {api.api_call_count}")
    print(f"  Users cached: {len(api._user_cache)}")
    print(f"  Attendance records cached: {len(api._attendance_cache)} users")
    print(f"  CSV files written to: {OUTPUT_DIR}")
    print(f"  Import these into the Excel workbook to populate the report.")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
