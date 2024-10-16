# Set-Signatures

This PowerShell CmdLet helps Office 365 administrators configure email signatures for users across the organization or for specific users in an automated manner.

## Features

**Connection to Microsoft 365**  
The script connects to the Office 365 tenant using Microsoft Graph and Exchange Online.

**Signature Configuration:**  
You can apply a custom HTML signature to all users in the organization, to specific users, or to groups of users.

**Dynamic Signatures:**  
The signature is generated dynamically using an HTML template and populated with user information such as name, job title, email address, and phone number.
OWA and Outlook Desktop Compatibility: The script configures signatures for both Outlook Web App (OWA) and Outlook Desktop App.

## Usage

The script can be easily executed through a batch file, which provides two options: applying the signature to all users or only to a specific user.

## CmdLet Parameters

| Parameter             | Action                                                                             |
| --------------------- | ---------------------------------------------------------------------------------- |
| **NombredeNegocio**   | Company name to include in the signature.                                          |
| **PathPlantillaHtml** | Path to the HTML template file used to generate the signature.                     |
| **CorreoUsuario**     | User's email address (optional). If omitted, signatures will be set for all users. |

## Example Usage

To apply a signature to _all users:_

```bash
PowerShell -ExecutionPolicy Bypass -File Set-Signatures.ps1 -NombredeNegocio "MyCompany" -PathPlantillaHtml "C:\Template\signature.html"
```

To apply a signature to a specific user:

```bash
PowerShell -ExecutionPolicy Bypass -File Set-Signatures.ps1 -NombredeNegocio "MyCompany" -P
```
