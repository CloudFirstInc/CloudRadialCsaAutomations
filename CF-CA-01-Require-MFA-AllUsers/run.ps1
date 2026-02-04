{
  "displayName": "CF-CA-01-Require-MFA-AllUsers",
  "state": "enabledForReportingButNotEnforced",
  "conditions": {
    "users": {
      "includeUsers": ["All"],
      "excludeUsers": ["{BREAK_GLASS_USER_OBJECT_ID_1}", "{BREAK_GLASS_USER_OBJECT_ID_2}"]
    },
    "applications": {
      "includeApplications": ["All"]
    },
    "clientAppTypes": ["all"]
  },
  "grantControls": {
    "operator": "OR",
    "builtInControls": ["mfa"]
  }
}