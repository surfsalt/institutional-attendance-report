#!/usr/bin/env python3
"""
Quick test: Can we resolve department names from the hierarchy API?

Tests THREE things:
  A) List ALL top-level hierarchy nodes (to find the correct root)
  B) Top-down: Get children of 'All Departments' node (the old approach)
  C) Reverse: For each course, ask 'what nodes is this course in?'
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


def api_get_safe(session, base_url, token, path):
    """GET that returns None on error instead of raising."""
    try:
        return api_get(session, base_url, token, path)
    except Exception as e:
        print(f"    FAILED: {e}")
        return None


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

    # ── Test A: List ALL hierarchy nodes ─────────────────────────────────
    print("=" * 60)
    print("TEST A: List ALL hierarchy nodes (find the root)")
    print("=" * 60)
    data = api_get_safe(session, base_url, token,
                        "/learn/api/public/v1/institutionalHierarchy/nodes")
    if data:
        nodes = data.get("results", [])
        print(f"  Total nodes found: {len(nodes)}")
        for n in nodes:
            parent = n.get("parentId", "NONE (root)")
            print(f"    id: {n.get('id', 'N/A')}")
            print(f"    title: {n.get('title', n.get('name', 'N/A'))}")
            print(f"    externalId: {n.get('externalId', 'N/A')}")
            print(f"    parentId: {parent}")
            print()
        
        # Check paging
        paging = data.get("paging", {})
        if paging.get("nextPage"):
            print(f"  (More pages available: {paging['nextPage']})")
            # Fetch next pages
            next_page = paging["nextPage"]
            while next_page:
                more = api_get_safe(session, base_url, token, next_page)
                if more:
                    extra = more.get("results", [])
                    for n in extra:
                        parent = n.get("parentId", "NONE (root)")
                        print(f"    id: {n.get('id', 'N/A')}")
                        print(f"    title: {n.get('title', n.get('name', 'N/A'))}")
                        print(f"    externalId: {n.get('externalId', 'N/A')}")
                        print(f"    parentId: {parent}")
                        print()
                    nodes.extend(extra)
                    next_page = more.get("paging", {}).get("nextPage")
                else:
                    break
            print(f"  Total nodes (all pages): {len(nodes)}")
    else:
        print("  Could not list nodes.")

    # ── Test B: Top-down (current approach) ──────────────────────────────
    print(f"\n{'=' * 60}")
    print(f"TEST B: Top-down — children of node {ALL_DEPARTMENTS_NODE_ID}")
    print("=" * 60)
    data = api_get_safe(session, base_url, token,
                        f"/learn/api/public/v1/institutionalHierarchy/nodes/{ALL_DEPARTMENTS_NODE_ID}/children")
    if data:
        children = data.get("results", [])
        print(f"  Children found: {len(children)}")
        for c in children[:5]:
            print(f"    - {c.get('title', c.get('name', 'N/A'))} (id: {c.get('id', 'N/A')})")
    else:
        print(f"  Could not fetch children (404 = node ID doesn't exist)")

    # ── Test C: Reverse lookup (new approach) ────────────────────────────
    print(f"\n{'=' * 60}")
    print("TEST C: Reverse — get hierarchy nodes FOR a course")
    print("=" * 60)

    # Grab a few courses from the most recent term
    print("  Fetching terms...")
    terms_data = api_get(session, base_url, token, "/learn/api/public/v1/terms")
    terms = terms_data.get("results", [])
    terms.reverse()
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

        # Try with expand=node
        data = api_get_safe(session, base_url, token,
                            f"/learn/api/public/v1/courses/{cid}/nodes?expand=node")
        if data:
            nodes = data.get("results", [])
            print(f"    Hierarchy nodes: {len(nodes)}")
            for n in nodes:
                node_detail = n.get("node", {})
                print(f"      nodeId: {n.get('nodeId', 'N/A')}")
                print(f"      title: {node_detail.get('title', 'N/A')}")
                print(f"      externalId: {node_detail.get('externalId', 'N/A')}")
                print(f"      parentId: {node_detail.get('parentId', 'N/A')}")
                print(f"      isPrimary: {n.get('isPrimary', 'N/A')}")
            if not nodes:
                # Also try without expand
                data2 = api_get_safe(session, base_url, token,
                                     f"/learn/api/public/v1/courses/{cid}/nodes")
                if data2:
                    nodes2 = data2.get("results", [])
                    print(f"    Without expand: {len(nodes2)} nodes")
                    for n in nodes2:
                        print(f"      raw: {json.dumps(n)}")
        print()

    print("=" * 60)
    print("DONE")
    print("=" * 60)


if __name__ == "__main__":
    main()
