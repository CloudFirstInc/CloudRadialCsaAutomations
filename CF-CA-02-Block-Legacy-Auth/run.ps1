{
  "displayName": "CF-CA-02-Block-Legacy-Auth",
  "state": "enabledForReportingButNotEnforced",
  "conditions": {
    "users": {
      "includeUsers": ["All"],
      "excludeUsers": ["{BREAK_GLASS_USER_OBJECT_ID_1}", "{BREAK_GLASS_USER_OBJECT_ID_2}"]
    },
    "applications": {
      "includeApplications": ["All"]
    },
    "clientAppTypes": ["other"]
  },
  "grantControls": {
    "operator": "OR",
    "builtInControls": ["block"]
  }
}