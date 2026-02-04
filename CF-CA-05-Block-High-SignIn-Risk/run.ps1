{
  "displayName": "CF-CA-05-Block-High-SignIn-Risk",
  "state": "enabledForReportingButNotEnforced",
  "conditions": {
    "users": {
      "includeUsers": ["All"],
      "excludeUsers": ["{BREAK_GLASS_USER_OBJECT_ID_1}", "{BREAK_GLASS_USER_OBJECT_ID_2}"]
    },
    "applications": {
      "includeApplications": ["All"]
    },
    "signInRiskLevels": ["high"]
  },
  "grantControls": {
    "operator": "OR",
    "builtInControls": ["block"]
  }
}