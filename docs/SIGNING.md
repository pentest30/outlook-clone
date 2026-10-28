# Signing the installer (Azure Trusted Signing)

Unsigned installers trigger Chrome's "dangerous file" warning and Windows SmartScreen. The build signs the
app's files and `OutlookClone.Desktop-win-Setup.exe` with **Azure Trusted Signing** (now also called *Azure
Artifact Signing*) when the repository has the secrets below; without them it still builds, unsigned, with a
warning.

Signing alone does not silence every warning on day one: SmartScreen and Chrome also look at how often a file
has been downloaded. Signed releases build that reputation; publish real versions (tags `v*`) rather than
handing out CI builds.

## 1. Azure (one time)

1. **Subscription and provider.** In an Azure subscription, register the resource provider
   `Microsoft.CodeSigning` (Subscription › Resource providers).
2. **Trusted Signing account.** Create one (Basic tier is enough). Note its **name** and its **region**; the
   endpoint is the region's, for example `https://weu.codesigning.azure.net/` for West Europe,
   `https://eus.codesigning.azure.net/` for East US.
3. **Identity validation.** In the account, *Identity validations* › *New*: **Organization** (company
   details, can take a few days) or **Individual**, depending on what is open in your country. Microsoft
   checks the identity before any certificate is issued.
4. **Certificate profile.** Once validated, *Certificate profiles* › *Create* › **Public Trust**, using that
   validation. Note the profile **name**.
5. **An identity for GitHub.** Microsoft Entra ID › *App registrations* › *New registration* (e.g.
   `outlook-clone-signing`), then *Certificates & secrets* › *New client secret*. Note the **tenant ID**,
   **client (application) ID** and the **secret value**.
6. **Permission to sign.** On the Trusted Signing account (or just the certificate profile), *Access control
   (IAM)* › *Add role assignment* › **Trusted Signing Certificate Profile Signer** › the app registration.

## 2. GitHub secrets

Repository › *Settings* › *Secrets and variables* › *Actions* › *New repository secret*:

| Secret | Value |
|---|---|
| `AZURE_TENANT_ID` | Directory (tenant) ID |
| `AZURE_CLIENT_ID` | Application (client) ID of the app registration |
| `AZURE_CLIENT_SECRET` | The client secret's value |
| `AZURE_SIGNING_ENDPOINT` | e.g. `https://weu.codesigning.azure.net/` |
| `AZURE_SIGNING_ACCOUNT` | Trusted Signing account name |
| `AZURE_SIGNING_PROFILE` | Certificate profile name |

The next build signs, then the *Check signature* step fails the run if `Setup.exe` is not validly signed.
Client secrets expire (at most 2 years): renew it in Entra ID and update `AZURE_CLIENT_SECRET`.

## Signing a local build

```powershell
# metadata.json
# { "Endpoint": "https://weu.codesigning.azure.net/",
#   "CodeSigningAccountName": "<account>", "CertificateProfileName": "<profile>" }

az login                      # or set AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET
pwsh build/package.ps1 -AzureSigningMetadata metadata.json
Get-AuthenticodeSignature releases/OutlookClone.Desktop-win-Setup.exe
```
