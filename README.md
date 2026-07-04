# Microsoft SharePoint File Sync for Nextcloud

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)

Recursively download a **SharePoint / OneDrive folder** (and everything below it) into a
Nextcloud **Group Folder**, preserving the full folder hierarchy. Authentication uses a single
Microsoft account via **delegated OAuth2** — no changes are required on the remote tenant,
as long as that account already has access to the shared folders.

> Built by [IANUSTEC](https://ianustec.com) as part of our open-source initiative to help
> organisations migrate from Microsoft to self-hosted infrastructure.

## Features

- Delegated OAuth2 (authorization-code flow) with a single service account; the refresh token
  is stored **encrypted** and access tokens are refreshed automatically.
- Recursive download of a SharePoint / OneDrive folder and all its sub-folders.
- Destination inside a Nextcloud **Group Folder** (team folder), with the SharePoint path
  replicated **idempotently** — existing folders are reused, never duplicated.
- Change detection by file size and modification time — unchanged files are skipped on re-runs.
- Manual **"Download now"** trigger from the admin panel with a live summary report (files
  downloaded / skipped / folders / errors).
- Uses Nextcloud's built-in HTTP client — **no external Composer runtime dependencies**.

## Requirements

- Nextcloud 27–32, PHP ≥ 8.1 (`curl`, `json` extensions).
- The [Group Folders](https://apps.nextcloud.com/apps/groupfolders) app installed and at least
  one group folder configured.
- A Microsoft work or school account that has access to the SharePoint folders you want to
  import (guest/external access is supported).

## How it works

```
Admin clicks "Download now"
        → refresh access token (delegated OAuth2)
        → resolve the SharePoint URL via the Microsoft Graph /shares endpoint
        → walk /drives/{id}/items/{id}/children recursively (paginated)
        → recreate each folder under the chosen Group Folder (idempotent)
        → download each new/changed file and write it via the Nextcloud Files API
```

Only the **Nextcloud filesystem adapter** and the **Microsoft Graph client** are specific to
this app; the rest is a thin, reusable sync skeleton that can be extended for other sources.

## Setup

### 1. Register an Azure AD application

1. Open **Microsoft Entra ID → App registrations → New registration**.
2. Account type: choose *Accounts in this organizational directory only* (single tenant) or
   *Accounts in any organizational directory* (multi-tenant), depending on where the service
   account lives.
3. Add a **Web** redirect URI matching the one shown on the app's admin page, e.g.
   `https://<your-nextcloud>/apps/neura_microsoft_sharepoint_sync/oauth/callback`.
4. Go to **API permissions → Add a permission → Microsoft Graph → Delegated permissions** and
   add:
   - `Files.Read.All`
   - `Sites.Read.All`
   - `User.Read`
   - `offline_access`

   Then click **Grant admin consent for \<your organisation\>**.
5. Go to **Certificates & secrets → New client secret** and copy the secret **value** (shown
   only once).

### 2. Configure the app in Nextcloud

Open **Administration settings → Microsoft SharePoint File Sync** and:

1. Fill in **Directory (tenant) ID**, **Application (client) ID** and the **client secret**,
   then click **Save settings**.
2. Click **Connect Microsoft** and sign in with the service account.
3. Paste the **SharePoint folder URL** (the browser URL of the folder you want to import).
4. Choose the destination **Group Folder**, an optional **sub-path**, and the **Nextcloud user
   to write as** (must belong to a group with write access on that group folder).
5. Click **Download now**.

Settings are persisted on each "Download now" click — no separate save is required for the
source/destination fields.

### Path mapping

With *Replicate the full SharePoint path* **enabled** (default), a source URL like:

```
https://contoso.sharepoint.com/sites/MARKETING/Shared%20Documents/Campaigns/2026/Q1-Report
```

is recreated under the destination as:

```
<GroupFolder>/<sub-path>/Campaigns/2026/Q1-Report/
    file1.pdf
    file2.xlsx
    subfolder/
        file3.docx
```

With the option **disabled**, only the leaf folder (`Q1-Report`) is created directly under
`<GroupFolder>/<sub-path>/`, omitting the intermediate path segments.

Folders that already exist are reused — re-running the download never creates duplicate trees.

## Installation

### From a GitHub release (recommended)

Download the latest `.tar.gz` from the
[Releases](https://github.com/ianustec/nextcloud-microsoft-sharepoint-sync/releases) page,
extract it into your Nextcloud `custom_apps/` directory, then enable the app:

```bash
cd /var/www/html/custom_apps
tar xzf /path/to/neura_microsoft_sharepoint_sync-<version>.tar.gz
php /var/www/html/occ app:enable neura_microsoft_sharepoint_sync
```

### From source (git clone)

```bash
cd /var/www/html/custom_apps
git clone https://github.com/ianustec/nextcloud-microsoft-sharepoint-sync.git neura_microsoft_sharepoint_sync
cd neura_microsoft_sharepoint_sync
# optional — generates the optimised PSR-4 autoloader; the app works without it too
composer install --no-dev --optimize-autoloader
php /var/www/html/occ app:enable neura_microsoft_sharepoint_sync
```

### Deploy script (Kubernetes / Docker / bare-metal)

A convenience `deploy.sh` is included. It builds the vendor autoloader, packages
the app, copies it into the target environment and runs `occ app:enable` automatically.

```bash
# Kubernetes — second arg is the Deployment name, not the pod name
./deploy.sh k8s <namespace> <deployment-name>
# e.g. ./deploy.sh k8s my-namespace nextcloud

# Docker / Docker Compose — pass the container name
./deploy.sh docker <container-name>
# e.g. ./deploy.sh docker nextcloud

# Bare metal — pass the Nextcloud root (directory that contains occ)
./deploy.sh local <nextcloud-root>
# e.g. ./deploy.sh local /var/www/html
```

## Security notes

- The client secret, refresh token and cached access token are encrypted at rest with
  Nextcloud's `ICrypto` (AES-256-CBC + HMAC).
- All admin API endpoints require administrator privileges (`@AdminRequired`).
- The OAuth callback is protected by a session `state` parameter against CSRF.
- The app uses Nextcloud's built-in HTTP client with **IPv4 forced** for all outbound
  connections (avoids broken IPv6 egress common in container environments).

## Roadmap

- Optional scheduled / incremental sync (background job with Microsoft Graph delta tokens).
- Client-credentials (app-only) flow for fully unattended deployments without a service account.
- Support for OneDrive personal accounts.

## Contributing

Contributions are welcome! Please read [CONTRIBUTING.md](CONTRIBUTING.md) before opening a
pull request.

## License

[AGPL-3.0-or-later](LICENSE)
