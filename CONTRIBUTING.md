# Contributing to Microsoft SharePoint File Sync for Nextcloud

Thank you for your interest in contributing! This document explains how to get
involved — from reporting bugs to submitting pull requests.

## Table of contents

- [Code of conduct](#code-of-conduct)
- [Reporting bugs](#reporting-bugs)
- [Requesting features](#requesting-features)
- [Development setup](#development-setup)
- [Submitting a pull request](#submitting-a-pull-request)
- [Coding standards](#coding-standards)
- [License](#license)

---

## Code of conduct

Be respectful and constructive. We follow the
[Contributor Covenant v2.1](https://www.contributor-covenant.org/version/2/1/code_of_conduct/).

---

## Reporting bugs

1. Search [existing issues](https://github.com/ianustec/nextcloud-microsoft-sharepoint-sync/issues)
   first to avoid duplicates.
2. Open a **Bug report** issue and fill in all sections of the template.
3. Include:
   - Nextcloud version and PHP version.
   - The exact error message (from the browser console **and** from
     `nextcloud.log` / `occ log:watch`).
   - Steps to reproduce.
   - What you expected to happen vs. what actually happened.

> **Never include real credentials, tokens, or personal data in issues.**

---

## Requesting features

Open a **Feature request** issue. Describe the use-case and why it cannot be
achieved with the current settings. The more concrete the proposal, the faster
we can evaluate it.

---

## Development setup

### Prerequisites

- PHP ≥ 8.1 with `curl`, `json`, `openssl` extensions.
- Composer ≥ 2.
- A local Nextcloud instance (Docker is the easiest path).
- The [Group Folders](https://apps.nextcloud.com/apps/groupfolders) app enabled.

### Clone and install

```bash
# 1. Fork the repository, then clone your fork
git clone https://github.com/<your-username>/nextcloud-microsoft-sharepoint-sync.git \
  neura_microsoft_sharepoint_sync

# 2. Symlink (or copy) the folder into Nextcloud's custom_apps/
ln -s "$(pwd)/neura_microsoft_sharepoint_sync" /var/www/html/custom_apps/

# 3. Generate the autoloader (no runtime deps, but needed for IDE completion)
cd neura_microsoft_sharepoint_sync
composer install

# 4. Enable the app
php /var/www/html/occ app:enable neura_microsoft_sharepoint_sync
```

### Project layout

```
appinfo/       App manifest (info.xml) and route definitions
lib/
  AppInfo/     Bootstrap entry points
  Controller/  HTTP controllers (admin settings, OAuth callback)
  Service/     Business logic (Microsoft Graph, Nextcloud Files, SyncEngine, Config)
  Settings/    Admin section / settings page registration
templates/     PHP admin UI template
js/            Vanilla JS for the admin page
css/           Styles for the admin page
img/           App icon
deploy.sh      Helper script for Kubernetes / Docker / bare-metal deployment
```

### Running lint & static analysis

```bash
# PHP syntax check
find lib appinfo templates -name '*.php' -print0 | xargs -0 php -l

# Optional: PHPStan (install separately)
composer require --dev phpstan/phpstan
vendor/bin/phpstan analyse lib --level=5
```

---

## Submitting a pull request

1. **Fork** the repository and create a feature branch off `main`:
   ```bash
   git checkout -b feat/my-feature
   ```
2. Make your changes. Keep commits focused and atomic.
3. Verify there are no PHP syntax errors (`php -l`).
4. Update `CHANGELOG.md` under an `## Unreleased` section.
5. Push your branch and open a pull request against `main`.
6. Fill in the PR template: describe what changed and how to test it.

### PR review checklist

- [ ] No client-specific, proprietary, or personal data in code or comments.
- [ ] New admin-only endpoints are decorated with `@AdminRequired`.
- [ ] Sensitive values (tokens, secrets) go through `ConfigService` and are
      encrypted at rest.
- [ ] No new Composer runtime dependencies introduced without prior discussion.
- [ ] `CHANGELOG.md` updated.

---

## Coding standards

- Follow **PSR-12** for PHP formatting.
- Class and interface names use `PascalCase`; methods and variables use
  `camelCase`; constants use `UPPER_SNAKE_CASE`.
- All public methods must have a PHPDoc block with `@param` and `@return` types.
- Prefer Nextcloud's built-in services (`IClientService`, `IConfig`,
  `ICrypto`, `ILogger`, …) over raw PHP functions.
- Keep `lib/Service/` classes single-responsibility; controllers should be thin.
- Do not add comments that just narrate what the code does — explain *why*,
  not *what*.

---

## License

By contributing you agree that your code will be released under the
[AGPL-3.0-or-later](LICENSE) license.
