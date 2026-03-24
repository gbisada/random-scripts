#!/usr/bin/env python3
"""
Microsoft 365 Reporting Hierarchy Tool

Authenticates as YOU (the logged-in user) via device code flow — no app
registration, client ID, or secrets required. Just run the script, open
the printed URL in your browser, enter the one-time code, and sign in
with your normal M365 / corporate credentials (MFA is fully supported).

The token is cached in ~/.m365_hierarchy_token so subsequent runs skip
the browser step until the token expires.

Requirements:
    pip install msal requests

Usage:
    python m365_reporting_hierarchy.py "alice@corp.com,bob@corp.com"
    python m365_reporting_hierarchy.py "alice@corp.com" --depth 2
    python m365_reporting_hierarchy.py "alice@corp.com" --output csv --out-file reports.csv
"""

import os
import sys
import csv
import argparse
import json
from collections import deque
from pathlib import Path

try:
    import msal
except ImportError:
    sys.exit("Error: 'msal' package not found. Run: pip install msal requests")

try:
    import requests
except ImportError:
    sys.exit("Error: 'requests' package not found. Run: pip install msal requests")


GRAPH_BASE = "<https://graph.microsoft.com/v1.0>"

# Microsoft Graph PowerShell SDK public client ID — pre-authorized for Graph
# API delegated access, no app registration or secret required.
PUBLIC_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"

# Delegated scopes needed. User.ReadBasic.All covers directReports lookups
# and is pre-consented for regular employees in most corporate tenants.
SCOPES = [
    "<https://graph.microsoft.com/User.Read>",
    "<https://graph.microsoft.com/User.ReadBasic.All>",
]

TOKEN_CACHE_PATH = Path.home() / ".m365_hierarchy_token"

USER_SELECT_FIELDS = ",".join([
    "id",
    "displayName",
    "mail",
    "userPrincipalName",
    "userType",
    "jobTitle",
    "department",
    "officeLocation",
    "accountEnabled",
    "onPremisesExtensionAttributes",
])


# ---------------------------------------------------------------------------
# Authentication — delegated, device code flow, no secrets
# ---------------------------------------------------------------------------

def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_PATH.exists():
        cache.deserialize(TOKEN_CACHE_PATH.read_text())
    return cache


def _save_cache(cache: msal.SerializableTokenCache) -> None:
    if cache.has_state_changed:
        TOKEN_CACHE_PATH.write_text(cache.serialize())
        TOKEN_CACHE_PATH.chmod(0o600)  # owner-read-only


def get_access_token(tenant: str = "organizations") -> str:
    """
    Obtain a delegated access token via device code flow.

    The user is prompted to visit <https://microsoft.com/devicelogin> and enter
    a short code. Supports MFA. Token is cached locally for reuse.

    Args:
        tenant: 'organizations' (any work/school account, default) or your
                specific tenant ID / domain if you want to pin to one tenant.
    """
    cache = _load_cache()
    authority = f"<https://login.microsoftonline.com/{tenant}>"
    app = msal.PublicClientApplication(
        PUBLIC_CLIENT_ID,
        authority=authority,
        token_cache=cache,
    )

    # Try silent refresh from cache first
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_cache(cache)
            return result["access_token"]

    # Fall back to interactive device code flow
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        sys.exit(f"Failed to create device flow: {flow.get('error_description', flow)}")

    print("\n" + "=" * 60)
    print("ACTION REQUIRED: Sign in to Microsoft 365")
    print("=" * 60)
    print(flow["message"])
    print("=" * 60 + "\n")

    result = app.acquire_token_by_device_flow(flow)  # blocks until user signs in
    if "access_token" not in result:
        error = result.get("error_description", result.get("error", "Unknown error"))
        sys.exit(f"Authentication failed: {error}")

    _save_cache(cache)
    print("Authentication successful.\n")
    return result["access_token"]


# ---------------------------------------------------------------------------
# Graph API helpers
# ---------------------------------------------------------------------------

def graph_get(session: requests.Session, url: str) -> dict | None:
    response = session.get(url)
    if response.status_code == 403:
        print(
            "\n[ERROR] Permission denied (403). Your tenant may require admin consent\n"
            "        for User.ReadBasic.All. Ask your IT admin to grant it, or try\n"
            "        running with --tenant <your-tenant-domain-or-id>.\n"
        )
        sys.exit(1)
    if response.status_code == 404:
        return None
    response.raise_for_status()
    return response.json()


def get_user(session: requests.Session, email: str) -> dict | None:
    url = f"{GRAPH_BASE}/users/{email}?$select={USER_SELECT_FIELDS}"
    return graph_get(session, url)


def get_direct_reports(session: requests.Session, user_id: str) -> list[dict]:
    reports = []
    url = (
        f"{GRAPH_BASE}/users/{user_id}/directReports"
        f"?$select={USER_SELECT_FIELDS}"
    )
    while url:
        data = graph_get(session, url)
        if not data:
            break
        reports.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return reports


# ---------------------------------------------------------------------------
# Core logic — BFS traversal of reporting hierarchy
# ---------------------------------------------------------------------------

def build_hierarchy(
    session: requests.Session,
    seed_emails: list[str],
    max_depth: int = 0,
) -> list[dict]:
    visited_ids: set[str] = set()
    results: list[dict] = []
    queue: deque[tuple[dict, int]] = deque()

    print(f"Resolving {len(seed_emails)} seed email(s)...", flush=True)
    for email in seed_emails:
        user = get_user(session, email.strip())
        if not user:
            print(f"  [WARN] Not found: {email}")
            continue
        uid = user.get("id")
        if uid and uid not in visited_ids:
            visited_ids.add(uid)
            queue.append((user, 0))
            print(f"  Found: {user.get('displayName')} <{email}>")

    print("\nTraversing reporting hierarchy...", flush=True)
    while queue:
        parent, depth = queue.popleft()
        reports = get_direct_reports(session, parent["id"])
        if not reports:
            continue

        indent = "  " * (depth + 1)
        print(f"{indent}{parent.get('displayName')} -> {len(reports)} direct report(s)")

        for report in reports:
            rid = report.get("id")
            if not rid or rid in visited_ids:
                continue
            visited_ids.add(rid)
            # Tag manager info — parent is already in memory, no extra API call
            report["_manager_name"]  = parent.get("displayName", "")
            report["_manager_email"] = parent.get("mail") or parent.get("userPrincipalName", "")
            results.append(report)
            next_depth = depth + 1
            if max_depth == 0 or next_depth < max_depth:
                queue.append((report, next_depth))

    return results


def format_record(user: dict) -> dict:
    odata_type = user.get("@odata.type", "")
    if "orgContact" in odata_type:
        user_type = "OrgContact"
    else:
        user_type = user.get("userType") or "Member"

    ext = user.get("onPremisesExtensionAttributes") or {}

    return {
        "displayName":     user.get("displayName", ""),
        "email":           user.get("mail") or user.get("userPrincipalName", ""),
        "type":            user_type,
        "jobTitle":        user.get("jobTitle", ""),
        "department":      user.get("department", ""),
        "officeLocation":  user.get("officeLocation", ""),
        "accountEnabled":  user.get("accountEnabled", ""),
        "lineManager":     user.get("_manager_name", ""),
        "lineManagerEmail": user.get("_manager_email", ""),
        "extAttr1":        ext.get("extensionAttribute1", ""),
        "extAttr2":        ext.get("extensionAttribute2", ""),
        "extAttr3":       ext.get("extensionAttribute3", ""),
        "extAttr4":       ext.get("extensionAttribute4", ""),
    }


# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------

COLUMNS = [
    "displayName", "email", "type", "jobTitle",
    "department", "officeLocation", "accountEnabled",
    "lineManager", "lineManagerEmail",
    "extAttr1", "extAttr2", "extAttr3", "extAttr4",
]


def output_table(records: list[dict]) -> None:
    if not records:
        print("\nNo reports found.")
        return
    widths = {col: len(col) for col in COLUMNS}
    for rec in records:
        for col in COLUMNS:
            widths[col] = max(widths[col], len(str(rec.get(col, ""))))
    header = "  ".join(col.ljust(widths[col]) for col in COLUMNS)
    sep    = "  ".join("-" * widths[col] for col in COLUMNS)
    print(f"\n{header}\n{sep}")
    for rec in records:
        print("  ".join(str(rec.get(col, "")).ljust(widths[col]) for col in COLUMNS))
    print(f"\nTotal: {len(records)} account(s)")


def output_csv(records: list[dict], filepath: str) -> None:
    with open(filepath, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=COLUMNS)
        writer.writeheader()
        writer.writerows(records)
    print(f"\nCSV written to: {filepath}  ({len(records)} record(s))")


def output_json(records: list[dict], filepath: str) -> None:
    with open(filepath, "w", encoding="utf-8") as fh:
        json.dump(records, fh, indent=2, ensure_ascii=False)
    print(f"\nJSON written to: {filepath}  ({len(records)} record(s))")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Drill down M365 org reporting lines — signs in as you, no app registration needed.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "emails",
        nargs="?",
        help="Comma-delimited seed email addresses. Prompted if omitted.",
    )
    parser.add_argument(
        "--tenant",
        default="organizations",
        metavar="TENANT",
        help=(
            "Your tenant domain or ID, e.g. <contoso.com> or the GUID "
            "(default: 'organizations' — accepts any work/school account)"
        ),
    )
    parser.add_argument(
        "--depth",
        type=int,
        default=0,
        metavar="N",
        help="Max depth to traverse (default: 0 = unlimited)",
    )
    parser.add_argument(
        "--output",
        choices=["table", "csv", "json"],
        default="table",
        help="Output format (default: table)",
    )
    parser.add_argument(
        "--out-file",
        default=None,
        metavar="FILE",
        help="Output file path (required for --output csv or json)",
    )
    parser.add_argument(
        "--clear-cache",
        action="store_true",
        help="Delete the cached token and force a fresh sign-in",
    )

    args = parser.parse_args()

    if args.clear_cache and TOKEN_CACHE_PATH.exists():
        TOKEN_CACHE_PATH.unlink()
        print("Token cache cleared.")

    if args.output in ("csv", "json") and not args.out_file:
        parser.error(f"--out-file is required when --output is '{args.output}'")

    raw_emails = args.emails or input("Enter comma-delimited email address(es): ")
    seed_emails = [e.strip() for e in raw_emails.split(",") if e.strip()]
    if not seed_emails:
        sys.exit("Error: No email addresses provided.")

    token = get_access_token(tenant=args.tenant)
    session = requests.Session()
    session.headers["Authorization"] = f"Bearer {token}"

    raw_results = build_hierarchy(session, seed_emails, max_depth=args.depth)
    records = [format_record(u) for u in raw_results]

    if args.output == "table":
        output_table(records)
    elif args.output == "csv":
        output_csv(records, args.out_file)
    elif args.output == "json":
        output_json(records, args.out_file)


if __name__ == "__main__":
    main()
