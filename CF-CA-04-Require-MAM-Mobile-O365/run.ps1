{
  "displayName": "CF-CA-04-Require-MAM-Mobile-O365",
  "state": "enabledForReportingButNotEnforced",
  "conditions": {
    "users": {
      "includeGroups": ["{MOBILE_BYOD_USERS_GROUP_OBJECT_ID}"]
    },
    "applications": {
      "includeApplications": ["Office365"]
    },
    "platforms": {
      "includePlatforms": ["iOS", "android"]
    },
    "clientAppTypes": ["mobileAppsAndDesktopClients"]
  },
  "grantControls": {
    "operator": "AND",
    "builtInControls": ["approvedApplication", "compliantApplication"]
  }
}