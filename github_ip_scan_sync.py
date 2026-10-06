#!/usr/bin/env python3
"""
GitHub Project Board Sync — Global IP Scan Kanban Board
Reads the 'Active Scans' sheet from the Excel file and syncs it into a
GitHub Projects (v2) board, creating issues and custom fields automatically.

Usage:
    python github_ip_scan_sync.py --token <GITHUB_PAT> --repo <owner/repo>
    python github_ip_scan_sync.py --token <GITHUB_PAT> --repo <owner/repo> --project-id <id>
    python github_ip_scan_sync.py --token <GITHUB_PAT> --repo <owner/repo> --dry-run
"""

import argparse
import json
import re
import sys
import time
import warnings
from datetime import datetime, date

import requests
import openpyxl

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

EXCEL_PATH = r"c:\Users\C5385241\Downloads\Kanban\Global IP Scan Kanban Board.xlsx"
SHEET_NAME = "Active Scans"

GITHUB_API = "https://api.github.com"
GITHUB_GRAPHQL = "https://api.github.com/graphql"

# Maps FOSS Scan State → GitHub project status option names
STATUS_MAP = {
    "Backlog": "Todo",
    "In Progress": "In Progress",
    "Done": "Done",
    "Cancelled": "Done",
}

# Priority label colours
PRIORITY_COLORS = {"P1": "d93f0b", "P2": "e4b000", "P3": "0075ca"}
LOCATION_COLORS = {
    "GTLC India": "5319e7",
    "GTLC Germany": "0052cc",
    "GTLC China": "b60205",
    "GTLC Vietnam": "006b75",
    "GLOBAL": "cccccc",
}

# ---------------------------------------------------------------------------
# Excel helpers
# ---------------------------------------------------------------------------

def load_active_scans(path: str) -> list[dict]:
    """Return list of dicts, one per non-empty data row."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[SHEET_NAME]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if not rows:
        return []

    headers = [str(h).strip() if h else f"col_{i}" for i, h in enumerate(rows[0])]
    records = []
    for row in rows[1:]:
        if not any(v is not None for v in row):
            continue
        record = {}
        for header, value in zip(headers, row):
            # Skip formula strings
            if isinstance(value, str) and value.startswith("="):
                value = None
            record[header] = value
        records.append(record)
    return records


def fmt_date(value) -> str | None:
    """Convert datetime/date to ISO date string."""
    if value is None:
        return None
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    return None


def build_issue_body(row: dict) -> str:
    """Build the markdown body for a GitHub issue from an Excel row."""
    def f(k):
        v = row.get(k)
        return str(v) if v is not None else "—"

    def fd(k):
        return fmt_date(row.get(k)) or "—"

    body = f"""## IP Scan Details

| Field | Value |
|---|---|
| **L1 Manager** | {f('L1 Manager')} |
| **L1 Development Unit** | {f('L1 Development Unit')} |
| **Dev Contact** | {f('Dev Contact')} |
| **Program Category** | {f('Program Category')} |
| **Delivery Mode** | {f('Delivery Mode')} |
| **Priority** | {f('Priority')} |
| **Severity** | {f('Severity')} |
| **Previously Scanned** | {f('Previously Scanned')} |
| **Scan Category** | {f('Scan Category')} |
| **Start of Development** | {fd('Start of Development')} |
| **End of Development** | {fd('End of Development')} |
| **RTC Date** | {fd('RTC')} |
| **Location** | {f('Location')} |
| **FOSS Analyst** | {f('FOSS Analyst')} |
| **FOSS Scan State** | {f('FOSS Scan State')} |
| **Black Duck Scan** | {f('Black Duck Scan')} |
| **FOSS Scan Assigned** | {fd('FOSS Scan Assigned')} |
| **FOSS Scan Closed** | {fd('FOSS Scan Closed')} |
| **Commercial Scan Status** | {f('Commercial Scan Status')} |
| **Commercial Analyst** | {f('Comm Analyst')} |
| **Web Service Scan Status** | {f('Web Service Scan Status')} |
| **Web Service Analyst** | {f('Web Service Analyst')} |
| **SLOC Count** | {f('SLOC Count')} |
| **Tracking Ticket Status** | {f('Tracking System Ticket Status')} |
| **RTC Passed** | {f('RTC Passed')} |
| **Owner #** | {f('Owner #')} |
| **Scan Submission Date** | {fd('Scan Submission Date')} |
| **Continuous Scan** | {f('Continuous scan')} |
| **Crown Jewel** | {f('Crown Jewel')} |
| **Closed within SLA** | {f('Closed within SLA')} |
| **Unreviewed Components** | {f('Unreviewed components')} |
| **RDM** | {fd('RDM')} |
| **Maintenance** | {f('Maintenance')} |
"""

    jira_foss = row.get("CTP Scan Ticket JIRA link")
    jira_ws = row.get("WS Scan Ticket JIRA link")
    comments = row.get("Comments")

    if jira_foss:
        body += f"\n**FOSS JIRA:** {jira_foss}\n"
    if jira_ws:
        body += f"\n**Web Service JIRA:** {jira_ws}\n"
    if comments:
        body += f"\n## Comments\n{comments}\n"

    return body.strip()


# ---------------------------------------------------------------------------
# GitHub REST helpers
# ---------------------------------------------------------------------------

class GitHubClient:
    def __init__(self, token: str, repo: str, dry_run: bool = False):
        self.token = token
        self.repo = repo          # "owner/repo"
        self.dry_run = dry_run
        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
        })
        # Cache
        self._existing_issues: dict[str, int] = {}   # title → number
        self._existing_labels: set[str] = set()
        self._rate_limit_remaining = 5000

    # ---- rate limiting -------------------------------------------------------

    def _check_rate(self, response: requests.Response):
        remaining = int(response.headers.get("X-RateLimit-Remaining", 1000))
        self._rate_limit_remaining = remaining
        if remaining < 10:
            reset = int(response.headers.get("X-RateLimit-Reset", time.time() + 60))
            wait = max(reset - int(time.time()), 1)
            print(f"  [rate-limit] Sleeping {wait}s to avoid hitting GitHub rate limit...")
            time.sleep(wait)

    def get(self, path: str, **kwargs) -> dict | list:
        url = f"{GITHUB_API}/{path}"
        resp = self.session.get(url, **kwargs)
        self._check_rate(resp)
        resp.raise_for_status()
        return resp.json()

    def post(self, path: str, payload: dict) -> dict:
        if self.dry_run:
            print(f"  [dry-run] POST /{path} {json.dumps(payload)[:120]}")
            return {}
        url = f"{GITHUB_API}/{path}"
        resp = self.session.post(url, json=payload)
        self._check_rate(resp)
        resp.raise_for_status()
        return resp.json()

    def patch(self, path: str, payload: dict) -> dict:
        if self.dry_run:
            print(f"  [dry-run] PATCH /{path} {json.dumps(payload)[:120]}")
            return {}
        url = f"{GITHUB_API}/{path}"
        resp = self.session.patch(url, json=payload)
        self._check_rate(resp)
        resp.raise_for_status()
        return resp.json()

    def graphql(self, query: str, variables: dict = None) -> dict:
        if self.dry_run:
            print(f"  [dry-run] GraphQL: {query[:80].replace(chr(10),' ')}")
            return {"data": {}}
        resp = self.session.post(
            GITHUB_GRAPHQL,
            json={"query": query, "variables": variables or {}},
        )
        self._check_rate(resp)
        resp.raise_for_status()
        result = resp.json()
        if "errors" in result:
            raise RuntimeError(f"GraphQL error: {result['errors']}")
        return result

    # ---- labels --------------------------------------------------------------

    def ensure_label(self, name: str, color: str = "ededed", description: str = "") -> None:
        if name in self._existing_labels:
            return
        if self.dry_run:
            print(f"  [dry-run] Ensure label: {name}")
            self._existing_labels.add(name)
            return
        # Fetch existing labels once
        if not self._existing_labels:
            page = 1
            while True:
                items = self.get(f"repos/{self.repo}/labels", params={"per_page": 100, "page": page})
                for item in items:
                    self._existing_labels.add(item["name"])
                if len(items) < 100:
                    break
                page += 1

        if name in self._existing_labels:
            return

        print(f"  Creating label: {name}")
        try:
            self.post(f"repos/{self.repo}/labels", {
                "name": name, "color": color, "description": description
            })
        except requests.HTTPError as e:
            if e.response.status_code == 422:
                pass  # already exists
            else:
                raise
        self._existing_labels.add(name)

    # ---- issues --------------------------------------------------------------

    def load_existing_issues(self) -> None:
        """Cache all existing issue titles to avoid duplicates."""
        print("Loading existing issues...")
        page = 1
        while True:
            items = self.get(
                f"repos/{self.repo}/issues",
                params={"state": "all", "per_page": 100, "page": page},
            )
            for item in items:
                if "pull_request" not in item:
                    self._existing_issues[item["title"]] = item["number"]
            if len(items) < 100:
                break
            page += 1
        print(f"  Found {len(self._existing_issues)} existing issues.")

    def create_or_get_issue(self, title: str, body: str, labels: list[str]) -> int | None:
        """Create a new issue or return existing issue number."""
        if title in self._existing_issues:
            return self._existing_issues[title]
        print(f"  Creating issue: {title[:80]}")
        data = self.post(
            f"repos/{self.repo}/issues",
            {"title": title, "body": body, "labels": labels},
        )
        number = data.get("number")
        if number:
            self._existing_issues[title] = number
        return number

    def close_issue(self, number: int) -> None:
        self.patch(f"repos/{self.repo}/issues/{number}", {"state": "closed"})

    # ---- GitHub Projects v2 --------------------------------------------------

    def get_repo_node_id(self) -> str:
        data = self.graphql("""
            query($owner: String!, $name: String!) {
              repository(owner: $owner, name: $name) { id }
            }
        """, {"owner": self.repo.split("/")[0], "name": self.repo.split("/")[1]})
        return data["data"]["repository"]["id"]

    def list_projects(self) -> list[dict]:
        owner, name = self.repo.split("/")
        data = self.graphql("""
            query($owner: String!, $name: String!) {
              repository(owner: $owner, name: $name) {
                projectsV2(first: 20) {
                  nodes { id number title }
                }
              }
            }
        """, {"owner": owner, "name": name})
        return data["data"]["repository"]["projectsV2"]["nodes"]

    def create_project(self, title: str, repo_node_id: str) -> dict:
        data = self.graphql("""
            mutation($repoId: ID!, $title: String!) {
              createProjectV2(input: { ownerId: $repoId, title: $title }) {
                projectV2 { id number title }
              }
            }
        """, {"repoId": repo_node_id, "title": title})
        return data["data"]["createProjectV2"]["projectV2"]

    def get_project_fields(self, project_id: str) -> dict:
        """Returns {fieldName: {id, type, options: {name: id}}}"""
        data = self.graphql("""
            query($projectId: ID!) {
              node(id: $projectId) {
                ... on ProjectV2 {
                  fields(first: 50) {
                    nodes {
                      ... on ProjectV2Field { id name dataType }
                      ... on ProjectV2SingleSelectField {
                        id name dataType
                        options { id name }
                      }
                      ... on ProjectV2IterationField { id name dataType }
                    }
                  }
                }
              }
            }
        """, {"projectId": project_id})
        fields = {}
        for node in data["data"]["node"]["fields"]["nodes"]:
            name = node.get("name", "")
            info = {"id": node["id"], "type": node.get("dataType", "TEXT"), "options": {}}
            for opt in node.get("options", []):
                info["options"][opt["name"]] = opt["id"]
            fields[name] = info
        return fields

    def create_project_field(self, project_id: str, name: str, data_type: str,
                              options: list[str] | None = None) -> dict:
        """Create a custom field. data_type: TEXT, NUMBER, DATE, SINGLE_SELECT."""
        if data_type == "SINGLE_SELECT":
            opts = [{"name": o, "color": "GRAY", "description": ""} for o in (options or [])]
            data = self.graphql("""
                mutation($projectId: ID!, $name: String!, $opts: [ProjectV2SingleSelectFieldOptionInput!]!) {
                  createProjectV2Field(input: {
                    projectId: $projectId
                    dataType: SINGLE_SELECT
                    name: $name
                    singleSelectOptions: $opts
                  }) { projectV2Field { ... on ProjectV2SingleSelectField { id name options { id name } } } }
                }
            """, {"projectId": project_id, "name": name, "opts": opts})
            node = data["data"]["createProjectV2Field"]["projectV2Field"]
            result = {"id": node["id"], "type": "SINGLE_SELECT", "options": {}}
            for opt in node.get("options", []):
                result["options"][opt["name"]] = opt["id"]
            return result
        else:
            data = self.graphql("""
                mutation($projectId: ID!, $name: String!, $dt: ProjectV2CustomFieldType!) {
                  createProjectV2Field(input: {
                    projectId: $projectId
                    dataType: $dt
                    name: $name
                  }) { projectV2Field { ... on ProjectV2Field { id name dataType } } }
                }
            """, {"projectId": project_id, "name": name, "dt": data_type})
            node = data["data"]["createProjectV2Field"]["projectV2Field"]
            return {"id": node["id"], "type": data_type, "options": {}}

    def add_issue_to_project(self, project_id: str, issue_node_id: str) -> str:
        """Returns the project item id."""
        data = self.graphql("""
            mutation($projectId: ID!, $contentId: ID!) {
              addProjectV2ItemById(input: { projectId: $projectId, contentId: $contentId }) {
                item { id }
              }
            }
        """, {"projectId": project_id, "contentId": issue_node_id})
        return data["data"]["addProjectV2ItemById"]["item"]["id"]

    def set_field_value(self, project_id: str, item_id: str, field_id: str,
                        field_type: str, value, option_id: str = None) -> None:
        if value is None:
            return
        if field_type == "SINGLE_SELECT":
            if option_id is None:
                return
            self.graphql("""
                mutation($projectId: ID!, $itemId: ID!, $fieldId: ID!, $optId: String!) {
                  updateProjectV2ItemFieldValue(input: {
                    projectId: $projectId itemId: $itemId fieldId: $fieldId
                    value: { singleSelectOptionId: $optId }
                  }) { projectV2Item { id } }
                }
            """, {"projectId": project_id, "itemId": item_id,
                  "fieldId": field_id, "optId": option_id})
        elif field_type == "TEXT":
            self.graphql("""
                mutation($projectId: ID!, $itemId: ID!, $fieldId: ID!, $val: String!) {
                  updateProjectV2ItemFieldValue(input: {
                    projectId: $projectId itemId: $itemId fieldId: $fieldId
                    value: { text: $val }
                  }) { projectV2Item { id } }
                }
            """, {"projectId": project_id, "itemId": item_id,
                  "fieldId": field_id, "val": str(value)})
        elif field_type == "NUMBER":
            try:
                num = float(value)
            except (TypeError, ValueError):
                return
            self.graphql("""
                mutation($projectId: ID!, $itemId: ID!, $fieldId: ID!, $val: Float!) {
                  updateProjectV2ItemFieldValue(input: {
                    projectId: $projectId itemId: $itemId fieldId: $fieldId
                    value: { number: $val }
                  }) { projectV2Item { id } }
                }
            """, {"projectId": project_id, "itemId": item_id,
                  "fieldId": field_id, "val": num})
        elif field_type == "DATE":
            date_str = fmt_date(value)
            if not date_str:
                return
            self.graphql("""
                mutation($projectId: ID!, $itemId: ID!, $fieldId: ID!, $val: Date!) {
                  updateProjectV2ItemFieldValue(input: {
                    projectId: $projectId itemId: $itemId fieldId: $fieldId
                    value: { date: $val }
                  }) { projectV2Item { id } }
                }
            """, {"projectId": project_id, "itemId": item_id,
                  "fieldId": field_id, "val": date_str})

    def get_issue_node_id(self, issue_number: int) -> str:
        owner, name = self.repo.split("/")
        data = self.graphql("""
            query($owner: String!, $name: String!, $num: Int!) {
              repository(owner: $owner, name: $name) {
                issue(number: $num) { id }
              }
            }
        """, {"owner": owner, "name": name, "num": issue_number})
        return data["data"]["repository"]["issue"]["id"]


# ---------------------------------------------------------------------------
# Sync logic
# ---------------------------------------------------------------------------

CUSTOM_FIELDS_DEF = [
    # (field_name, data_type, options_list_or_None)
    ("Priority",          "SINGLE_SELECT", ["P1", "P2", "P3"]),
    ("Location",          "SINGLE_SELECT", ["GTLC India", "GTLC Germany", "GTLC China", "GTLC Vietnam", "GLOBAL"]),
    ("Program Category",  "SINGLE_SELECT", ["On-Demand (SaaS)", "On-Mobile Device", "On-Premise", "Outbound OEM", "Source Code"]),
    ("FOSS Analyst",      "TEXT",          None),
    ("L1 Manager",        "TEXT",          None),
    ("End of Development","DATE",          None),
    ("RTC",               "DATE",          None),
    ("SLOC Count",        "NUMBER",        None),
    ("Black Duck Scan",   "SINGLE_SELECT", ["Y", "N"]),
    ("Continuous Scan",   "SINGLE_SELECT", ["Y", "N"]),
    ("Crown Jewel",       "SINGLE_SELECT", ["Y", "N"]),
    ("Severity",          "SINGLE_SELECT", ["S1", "S2", "S3"]),
    ("Owner #",           "TEXT",          None),
]


def ensure_project_fields(client: GitHubClient, project_id: str) -> dict:
    """Ensure all custom fields exist; return field map {name: field_info}."""
    print("Checking project fields...")
    existing = client.get_project_fields(project_id)
    fields = dict(existing)

    for name, dtype, options in CUSTOM_FIELDS_DEF:
        if name in fields:
            continue
        print(f"  Creating field: {name} ({dtype})")
        info = client.create_project_field(project_id, name, dtype, options)
        fields[name] = info
        time.sleep(0.3)

    return fields


def sync_row(client: GitHubClient, row: dict, project_id: str, fields: dict) -> None:
    program = row.get("Program (Delivery)")
    if not program:
        return

    title = str(program).strip()
    foss_state = str(row.get("FOSS Scan State") or "Backlog").strip()
    priority = str(row.get("Priority") or "").strip()
    location = str(row.get("Location") or "").strip()

    # Build labels
    labels = []
    if priority and priority in PRIORITY_COLORS:
        client.ensure_label(priority, PRIORITY_COLORS[priority], f"Priority {priority}")
        labels.append(priority)
    if location and location in LOCATION_COLORS:
        client.ensure_label(location, LOCATION_COLORS[location], f"Location: {location}")
        labels.append(location)
    prog_cat = str(row.get("Program Category") or "").strip()
    if prog_cat:
        client.ensure_label(prog_cat, "e99695", "Program category")
        labels.append(prog_cat)
    foss_label = f"FOSS:{foss_state}"
    client.ensure_label(foss_label, "f9d0c4", "FOSS scan state")
    labels.append(foss_label)

    # Create or retrieve issue
    body = build_issue_body(row)
    issue_number = client.create_or_get_issue(title, body, labels)
    if issue_number is None:
        return

    if client.dry_run:
        return

    # Close cancelled issues
    if foss_state == "Cancelled":
        client.close_issue(issue_number)

    # Add to project
    issue_node_id = client.get_issue_node_id(issue_number)
    item_id = client.add_issue_to_project(project_id, issue_node_id)

    # Set Status field
    status_field = fields.get("Status")
    if status_field:
        status_name = STATUS_MAP.get(foss_state, "Todo")
        opt_id = status_field["options"].get(status_name)
        client.set_field_value(project_id, item_id, status_field["id"],
                               "SINGLE_SELECT", status_name, opt_id)

    # Set custom fields
    field_values = {
        "Priority":           (row.get("Priority"), None),
        "Location":           (row.get("Location"), None),
        "Program Category":   (row.get("Program Category"), None),
        "FOSS Analyst":       (row.get("FOSS Analyst"), None),
        "L1 Manager":         (row.get("L1 Manager"), None),
        "End of Development": (row.get("End of Development"), None),
        "RTC":                (row.get("RTC"), None),
        "SLOC Count":         (row.get("SLOC Count"), None),
        "Black Duck Scan":    (row.get("Black Duck Scan"), None),
        "Continuous Scan":    (row.get("Continuous scan"), None),
        "Crown Jewel":        (row.get("Crown Jewel"), None),
        "Severity":           (row.get("Severity"), None),
        "Owner #":            (row.get("Owner #"), None),
    }

    for fname, (fval, _) in field_values.items():
        if fval is None:
            continue
        finfo = fields.get(fname)
        if not finfo:
            continue
        ftype = finfo["type"]
        opt_id = None
        if ftype == "SINGLE_SELECT":
            opt_id = finfo["options"].get(str(fval).strip())
            if not opt_id:
                continue
        client.set_field_value(project_id, item_id, finfo["id"], ftype, fval, opt_id)

    time.sleep(0.2)   # gentle pacing to avoid secondary rate limits


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def parse_args():
    parser = argparse.ArgumentParser(description="Sync Global IP Scan Excel → GitHub Project Board")
    parser.add_argument("--token",      required=True, help="GitHub Personal Access Token (needs repo + project scopes)")
    parser.add_argument("--repo",       required=True, help="owner/repo  e.g.  my-org/ip-scan-board")
    parser.add_argument("--excel",      default=EXCEL_PATH, help="Path to the Excel file")
    parser.add_argument("--project-id", help="Existing GitHub Projects v2 node ID (skip creation)")
    parser.add_argument("--project-title", default="Global IP Scan Kanban", help="Title for a new project")
    parser.add_argument("--dry-run",    action="store_true", help="Print what would happen without making API calls")
    parser.add_argument("--limit",      type=int, default=0, help="Process only first N rows (0 = all)")
    parser.add_argument("--state-filter", default="", help="Only sync rows where FOSS Scan State matches (e.g. 'In Progress')")
    return parser.parse_args()


def main():
    args = parse_args()
    client = GitHubClient(args.token, args.repo, dry_run=args.dry_run)

    # ------------------------------------------------------------------
    # 1. Load Excel data
    # ------------------------------------------------------------------
    print(f"\nLoading Excel: {args.excel}")
    records = load_active_scans(args.excel)
    print(f"  Loaded {len(records)} rows from '{SHEET_NAME}'")

    if args.state_filter:
        records = [r for r in records if str(r.get("FOSS Scan State") or "").strip() == args.state_filter]
        print(f"  After state filter '{args.state_filter}': {len(records)} rows")

    if args.limit:
        records = records[:args.limit]
        print(f"  Limiting to first {args.limit} rows")

    if not records:
        print("No records to process. Exiting.")
        sys.exit(0)

    # ------------------------------------------------------------------
    # 2. Resolve / create GitHub Project
    # ------------------------------------------------------------------
    if args.project_id:
        project_id = args.project_id
        print(f"\nUsing existing project: {project_id}")
    else:
        print("\nFetching repository node ID...")
        if not args.dry_run:
            repo_node_id = client.get_repo_node_id()
        else:
            repo_node_id = "DRY_RUN_REPO_ID"

        print(f"Checking for existing project '{args.project_title}'...")
        if not args.dry_run:
            projects = client.list_projects()
            match = next((p for p in projects if p["title"] == args.project_title), None)
            if match:
                project_id = match["id"]
                print(f"  Found existing project #{match['number']}: {match['title']}")
            else:
                print(f"  Creating new project: {args.project_title}")
                proj = client.create_project(args.project_title, repo_node_id)
                project_id = proj["id"]
                print(f"  Created project #{proj['number']}: {proj['title']}")
        else:
            project_id = "DRY_RUN_PROJECT_ID"

    # ------------------------------------------------------------------
    # 3. Ensure custom fields
    # ------------------------------------------------------------------
    print("\nEnsuring project fields...")
    if not args.dry_run:
        fields = ensure_project_fields(client, project_id)
    else:
        fields = {name: {"id": f"DRY_{name}", "type": dtype, "options": {}}
                  for name, dtype, _ in CUSTOM_FIELDS_DEF}
        fields["Status"] = {"id": "DRY_STATUS", "type": "SINGLE_SELECT",
                            "options": {"Todo": "OPT1", "In Progress": "OPT2", "Done": "OPT3"}}

    # ------------------------------------------------------------------
    # 4. Load existing issues (deduplication)
    # ------------------------------------------------------------------
    if not args.dry_run:
        client.load_existing_issues()

    # ------------------------------------------------------------------
    # 5. Sync each row
    # ------------------------------------------------------------------
    print(f"\nSyncing {len(records)} records...\n")
    for i, row in enumerate(records, 1):
        program = str(row.get("Program (Delivery)") or "").strip()
        state = str(row.get("FOSS Scan State") or "").strip()
        print(f"[{i}/{len(records)}] {program[:70]}  [{state}]")
        try:
            sync_row(client, row, project_id, fields)
        except Exception as exc:
            print(f"  ERROR: {exc}")

    print(f"\nDone. Processed {len(records)} records.")
    if args.dry_run:
        print("(dry-run mode — no changes were made)")


if __name__ == "__main__":
    main()
