# Synchronizing AzureAD Groups with Teams

## Synopsis

Script for synchronizing AzureAD groups with Microsoft Teams. <br>
Files:<br>
**Security_groups_Teams_sync.csv** – CSV file which represents relation between Security Groups in AzureAD with Teams and Channels in Microsoft Teams<br>
**cred.clixml** – file with securely stored credentials, needed to connect to AzureAD and Microsoft Teams.<br>
**Logs** – Logs from running script are generated in folder “Logs”. It contains information about time and added / removed users.

### CSV file example:

| Team (Teams) | Channel (Teams) | AAD Security group |
|--------------|-----------------|--------------------|
| admin        | Branch1         | admins_branch1     |
| admin        | Branch2         | admins_branch2     |
| Consulting   | Branch1         | Consulting_branch1 |
| Consulting   | Branch2         | Consulting_branch2 |