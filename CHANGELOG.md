# Changelog

All notable changes to this project are documented in this file.
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## Unreleased

## 1.0.2 – 2026-10-31

### Changed
- Compatible with Nextcloud 33 (`max-version` 33). Verified on Nextcloud 33.0.10: app enables, every class loads, all imported Nextcloud APIs exist. No code changes.
- `appinfo/info.xml` version aligned with the release (1.0.1 was tagged with 1.0.0 in info.xml).

## 1.0.1

### Fixed
- Force IPv4 for all outbound HTTP requests (`force_ip_resolve: v4`) to avoid
  silent failures on container networks where IPv6 egress is unavailable
  (cURL error 7).

## 1.0.0

- Initial release.
- Delegated OAuth2 (authorization-code flow) with a single Microsoft account; refresh token stored encrypted.
- Recursive download of a SharePoint / OneDrive folder and all its sub-folders.
- Destination inside a Nextcloud Group Folder, with the SharePoint path replicated idempotently.
- Change detection by size and modification time to skip unchanged files on re-runs.
- Manual "Download now" trigger from the admin panel with a live summary report.
