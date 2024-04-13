### Set-Signatures

Easily manage Outlook users signatures with automation.

This script has been developed to assist Microsoft Office 365 or Exchange administrators in configuring the users signatures.

Typically, as an Office 365 or Exchange admin you are tasked with the very repetitive job of configuring the users signatures. This involves including information such as user name, business email, job title, phone number, and other data in the signature.

To assist with this task, this script utilizes ExchangeOnlineManagement Powershell Module and Microsoft Graph to retrieve the user information from the organization's environment and generate the signatures based on an HTML template.

### Overview

With this script you can

- Set the desired custom signature.
- Set the same signature to all the users of organization with just one call to the script.
- Set the signature for a single user.
- Use different templates on different users.

### Use
Nowadays there are two options to run the script.

- Navigate to the script path with Powershell (Admin) and execute:
```
.\Set-Signatures.ps1 -TemplatePath <"Path of the signature"> -BusinessName <"Test Company"> [Optional] -UserMail <"test@company.net">
```
Or
- Navigate to the script folder and run ```Execute.bat```

### Planned Updates
- Specify a filter to determine which users receive the signature.
- Include custom fields of the user information.
- Set signatures on Desktop App.