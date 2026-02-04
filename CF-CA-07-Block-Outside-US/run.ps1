{
  "displayName": "CF-CA-07-Block-Outside-US",
  "state": "enabledForReportingButNotEnforced",
  "conditions": {
    "users": {
      "includeUsers": ["All"],
      "excludeUsers": ["{BREAK_GLASS_USER_OBJECT_ID_1}", "{BREAK_GLASS_USER_OBJECT_ID_2}"]
    },
    "applications": {
      "includeApplications": ["All"]
    },
    "locations": {
      "includeLocations": ["All"],
      "excludeLocations": ["{US_NAMED_LOCATION_ID}"]
    },
    "clientAppTypes": ["all"]
  },
  "grantControls": {
    "operator": "OR",
    "builtInControls": ["block"]
  }
}