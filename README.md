# Introduction

> _Note: Any scripts or source code in this solution has not been tested or developed in the intended target environment. Additional changes might be required._

The Banker Portal solution consists of 3 main folders:

- _\\provisioning\\_
- _\\teams\\_
- _\\document-portal\\_

Any directory paths provided in examples below have been removed for clarity and brevity.

## Provisioning

To provision the opportunities `\provisioning\Provisioning-Engine.ps1` is used. The script loops over all opportunities specified in `\provisioning\sites-and.members.csv` and

- Adding metadata to the property bag of each site such as client, opportunity id and solution constant (used to find all relevant opportunities with SharePoint Search)
- Creates the group `"Banker Portal Documents Members"` and assigns it Read permissions to the opportunity and Contribute permissions to the opportunity document library.

To run the script (_requires admin sign-in and PnP.PowerShell module_):

_(powershell)_
```powershell
cd \provisioning\
\provisioning> .\Provisioning-Engine.ps1
```

For future reference `\provisioning\reference` folder contains the provisioning script used internally for the POC.

## Teams Manifest

To support SharePoint access and sign-in from the Teams application line the `\teams\manifest.json` must be updated with **SharePoint Online Client Extensibility Web Application Principal** client ID which can be found in [Azure AD](https://aad.portal.azure.com#dashboard) under _Azure Active Directory > App registrations > All applications > Search > SharePoint Online Client Extensibility Web Application Principal_

### manifest.json (line 51)

```json
  "webApplicationInfo": {
    "resource": "https://ubscloudeng.sharepoint.com/",
    "id": "{SharePoint Online Client Extensibility Web Application Principal ID}"
  }
```

Once update with the corresponding GUID all files in the \_\teams\_ can be packaged to a .zip. The resulting .zip should not contain any sub-folders - only the files:

- color.png
- manifest.json
- outline.png


_(powershell)_
```powershell
cd \teams\
\teams> Compress-Archive "manifest.json","color.png","outline.png" manifest.zip
```

The resulting zip can be uploaded to the Teams organization app catalog.

## Banker Portal (SPFx)

To access the Banker Portal the SharePoint Framework (SPFx) packages must be deployed to the SharePoint app catalog. The Banker Portal WebPart will be the application bankers' access in Microsoft Teams via the Teams manifest (`.zip`) created in the previous step and displays the content of the sites provisioned in the first step (`\provisioning\Provisioning-Engine.ps1`) of this guide.

Used in demo environment:

- Node (version 10.20.1)
- NPM (version 7.7.6)
- Gulp (CLI version 2.2.0)

### Install packages

```cmd
cd \document-portal\
\document-portal> npm i
```

### Create app package

```cmd
cd \document-portal\
\document-portal> gulp bundle --ship
\document-portal> gulp package-solution --ship
```

The app package will be available at location `.\document-portal\sharepoint\solution\document-portal.sppkg` and can be uploaded to SharePoint app catalog.
