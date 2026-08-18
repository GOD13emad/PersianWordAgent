# PersianWordAgent

**PersianWordAgent** is a Windows desktop assistant for controlled formatting, auditing, and repair of Persian Microsoft Word documents.

> این مخزن برای انتشار نسخه‌های اجرایی PersianWordAgent است. کد منبع برنامه در این مخزن منتشر نمی‌شود.

## Current public release

**Version 1.0.0 — QH01 Accepted Preview**

- Release tag: `v1.0.0-qh01-preview`
- Public product version: `1.0.0`
- Windows file version: `1.0.0.0`
- Validated internal engine baseline: `2.0.0-m3.3-dev41`
- Distribution: Windows x64 binary only

The internal engine baseline identifies the accepted development engine used to build the public Version 1 package; it is not the public product version.

## Validated capabilities

- Persian/RTL document handling
- profile-based document classification and repair
- controlled-copy workflow (original input remains unchanged)
- Word-native Table of Contents refresh
- heading, table, caption, and pagination integrity gates
- fail-closed checks for unsupported or unsafe states
- final Microsoft Word rendering validation

## Distribution model

This is a **binary-release-only** repository.

- Source code is **not published** here.
- Official Windows binaries are distributed only through this repository's **Releases** page.
- Do not download PersianWordAgent from unofficial mirrors.
- See `LICENSE.txt` for permitted use and restrictions.

## Requirements

- Windows x64
- Microsoft Word desktop for Word-native document operations

## Safety model

PersianWordAgent is designed around controlled copies and explicit validation gates. The application should not be treated as a substitute for retaining backups of important documents.

## Privacy

Do not attach confidential Word documents, private evidence packages, credentials, access tokens, or organization-internal files to public GitHub issues.

## Release verification

Every official release includes a `SHA256SUMS.txt` file. Verify the SHA-256 hash of downloaded packages before use.

Current Windows x64 Version 1 package SHA-256:

`0DFF8B2FAD92DE03DB91C33BDCC8FCA6914ECF1B4B8D41694952D4ECED336906`

## Project status

- QH01: **ACCEPTED**
- Public release: **Version 1.0.0 preview**
- Next development phase: **FD01**

## License

PersianWordAgent is **not open source**. Binary use is governed by the **PersianWordAgent Binary Distribution License 1.0** in `LICENSE.txt`.

Copyright © 2026 GOD13emad. All rights reserved.
