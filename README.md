## TempoOutlookSync

- This application requires to have Classic Outlook to be installed **and preferably running**, because it uses COM objects to avoid the MS Graph API.
- This application requires you to provide your Jira UserId, Email and ApiToken as well as your Tempo API Token once.

**The application runs in the background and syncs your Capacity Planner Times into the Outlook Calendar.**

Your `application.toml` should look like this:
```toml
# Text with '#' at the start can be ignored as they are just comments

Email = "your.email@provider.de"
JiraApiToken = "your-super-long-secret-jira-api-token"
UserId = "some-numbers-this-time-b64"
TempoApiToken = "your-secret-tempo-api-token"
```

### Help for getting the required application.toml values:
- How to create a Tempo API Token: https://help.tempo.io/timesheets/latest/using-rest-api-integrations
- How to get my Jira UserId: https://www.storylane.io/tutorials/how-to-find-user-id-in-jira
- How to create a Jira API Token: https://docs.adaptavist.com/w4j/latest/quick-configuration-guide/add-sources/how-to-generate-jira-api-token
- How to get my Jira Email: Just click your profile picture at the top right in Jira
