# random-scripts

A collection of utility scripts.

## Scripts

### m365-reports.py

**Microsoft 365 Reporting Hierarchy Tool**

Retrieves a user's reporting hierarchy from Microsoft 365 / Azure AD using the Microsoft Graph API. Authenticates via device code flow — no app registration, client ID, or secrets required.

**Features:**
- Device code flow authentication (supports MFA)
- Token caching in `~/.m365_hierarchy_token` for subsequent runs
- Configurable traversal depth
- Multiple output formats: text, JSON, CSV

**Requirements:**
```
pip install msal requests
```

**Usage:**
```
python m365-reports.py "alice@corp.com,bob@corp.com"
python m365-reports.py "alice@corp.com" --depth 2
python m365-reports.py "alice@corp.com" --output csv --out-file reports.csv
```