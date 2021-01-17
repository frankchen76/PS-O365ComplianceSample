# PS-O365ComplianceSample
A PS sample for O365 Security &amp; Compliance PS

## AddRetentionUsers.ps1
A PS sample to add large amount of users to retention policy due to 1000 user limit for a retention policy. 

### Usage
configure the following parameters inside the AddRetentionUsers.ps1: 
* $userLimit: the user limit is 1000 users. For testing, you can change this number to a small number for a testing. 
* pnPrefix: the retention policy prefix name. The PS creates a retention policy name by appending a sequence number after prefix and uses the name to create a retention policy. 
* credName: the credential name which PS uses to connect to EXO. the credential needs to be created from windows credential store. 
* replace Get-GermanyUsers function to retrieve the users based on criteria. The current logic retrieves first 10 users created after 2019-10-01. You can combine the officeLocation or other properties to determine the users

After you confire above parameters, you can run the following PS to start the process: 

```PowerShell
.\AddRetentionUsers.ps1
```
