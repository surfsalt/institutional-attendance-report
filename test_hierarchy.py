#!/usr/bin/env python3
"""
Quick test: Can we resolve department names from the hierarchy API?

This script authenticates, grabs a few courses from the selected term,
and tries TWO approaches to see which one works:

  A) TOP-DOWN: Get children of the 'All Departments' node
  B) REVERSE:  For each course, ask 'what nodes is this course in?'

Run this BEFORE updating extract.py so we know which approach to use.
"""

import configparser
import os
import sys
import json

try:
    import requests
except ImportError:
    print("ERROR: 'requests' not installed. Run: pip install requests")
    sys.exit(1)

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "extract_config.ini")
ALL_DEPARTMENTS_NODE_ID = "05d8bd91-8efb-476c-91b4-98138168afab"


def load_config():
    config = configparser.ConfigParser()
    config.read(CONFIG_FILE)
    return config


def api_get(session, base_url, token, path):
    """Simple authenticated GET — no retry, no paging."""
    url = f"{base_url}{path}"
    resp = session.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    resp.raise_for_status()
    return resp.json()


def main():
    config = load_config()
    base_url = config["blackboard"]["base_url"].rstrip("/")
    api_key = config["blackboard"]["api_key"]
    api_secret = config["blackboard"]["api_secret"]

    session = requests.Session()

    # Authenticate
    print("[Auth] Getting token...")
    resp = session.post(
        f"{base_url}/learn/api/public/v1/oauth2/token",
        data={"grant_type": "client_credentials"},
        auth=(api_key, api_secret),
        timeout=15,
    )
    resp.raise_for_status()
    token = resp.json()["access_token"]
    print(f"[Auth] OK\n")

    # ── Test A: Top-down (current approach) ─────────────────────────────
    print("=" * 60)
    print("TEST A: Top-down — get children of 'All Departments' node")
    print("=" * 60)
    try:
        data = api_get(session, base_url, token,
                       f"/learn/api/public/v1/institutionalHierarchy/nodes/{ALL_DEPARTMENTS_NODE_ID}/children")
        children = data.get("results", [])
        print(f"  Children found: {len(children)}")
        for c in children[:5]:
            print(f"    - {c.get('name', c.get('title', 'N/A'))} (id: {c.get('id', 'N/A')})")
        if len(children) > 5:
            print(f"    ... and {len(children) - 5} more")
        if not children:
            print("  ⚠ No children found — this is why the department column is blank!")
            print(f"    Node ID used: {ALL_DEPARTMENTS_NODE_ID}")
            print(f"    Raw response: {json.dumps(data, indent=2)[:500]}")
    except Exception as e:
        print(f"  ✗ FAILED: {e}")

    # If children were found, try getting courses for the first one
    if children:
        first_node = children[0]
        nid = first_node.get("id", "")
        print(f"\n  Fetching courses for first department node '{first_node.get('name', 'N/A')}'...")
        try:
            data = api_get(session, base_url, token,
                           f"/learn/api/public/v1/institutionalHierarchy/nodes/{nid}/courses")
            courses = data.get("results", [])
            print(f"  Courses found: {len(courses)}")
            for c in courses[:3]:
                print(f"    - courseId: {c.get('courseId', 'N/A')}, isPrimary: {c.get('isPrimary', 'N/A')}")
        except Exception as e:
            print(f"  ✗ FAILED: {e}")

    # ── Test B: Reverse lookup (new approach) ───────────────────────────
    print(f"\n{'=' * 60}")
    print("TEST B: Reverse — get hierarchy nodes FOR a course")
    print("=" * 60)

    # First grab a few courses from the most recent term
    print("  Fetching terms...")
    terms_data = api_get(session, base_url, token, "/learn/api/public/v1/terms")
    terms = terms_data.get("results", [])
    terms.reverse()  # newest first
    print(f"  Found {len(terms)} terms, using: {terms[0].get('name', 'N/A')}")

    term_id = terms[0]["id"]
    courses_data = api_get(session, base_url, token,
                           f"/learn/api/public/v3/courses?termId={term_id}&limit=5")
    sample_courses = courses_data.get("results", [])
    print(f"  Sample courses: {len(sample_courses)}\n")

    for course in sample_courses:
        cid = course.get("id")
        ext_id = course.get("courseId", cid)
        name = course.get("name", "Unknown")
        print(f"  Course: {ext_id} — {name}")
        print(f"    Internal ID: {cid}")

        try:
            # Try with expand=node to get the full node details
            data = api_get(session, base_url, token,
                           f"/learn/api/public/v1/courses/{cid}/nodes?expand=node")
            nodes = data.get("results", [])
            print(f"    Hierarchy nodes: {len(nodes)}")
            for n in nodes:
                node_detail = n.get("node", {})
                print(f"      - nodeId: {n.get('nodeId', 'N/A')}")
                print(f"        title: {node_detail.get('title', 'N/A')}")
                print(f"        externalId: {node_detail.get('externalId', 'N/A')}")
                print(f"        parentId: {node_detail.get('parentId', 'N/A')}")
                print(f"        isPrimary: {n.get('isPrimary', 'N/A')}")
                if node_detail.get("parentId") == ALL_DEPARTMENTS_NODE_ID:
                    print(f"        ✓ This IS a top-level department!")
            if not nodes:
                print(f"      (no hierarchy nodes — course not assigned to any department)")
        except Exception as e:
            print(f"    ✗ FAILED: {e}")

        print()

    print("=" * 60)
    print("DONE — check above to see which approach returns department data.")
    print("=" * 60)


if __name__ == "__main__":
    main()
