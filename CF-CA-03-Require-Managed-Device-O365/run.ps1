{
  "displayName": "CF-CA-03-Require-Managed-Device-O365",
  "state": "enabledForReportingButNotEnforced",
  "conditions": {
    "users": {
      "includeGroups": ["{CORP_USERS_GROUP_OBJECT_ID}"],
      "excludeUsers": ["{BREAK_GLASS_USER_OBJECT_ID_1}", "{BREAK_GLASS_USER_OBJECT_ID_2}"]
    },
    "applications": {
      "includeApplications": ["Office365"]
    },
    "clientAppTypes": ["all"]
  },
  "grantControls": {
    "operator": "OR",
    "builtInControls": ["compliantDevice", "domainJoinedDevice"]
  }
}