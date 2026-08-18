# PersianWordAgent

**PersianWordAgent** is a Windows desktop assistant for controlled formatting, auditing, and repair of Persian Microsoft Word documents.

> این مخزن برای انتشار نسخه‌های اجرایی PersianWordAgent است. کد منبع برنامه در این مخزن منتشر نمی‌شود.

## Current public milestone

**QH01 Accepted Preview — dev41**

The accepted QH01 build has completed controlled technical, Microsoft Word pagination, TOC freshness, and final desktop GUI validation.

Key validated behaviors include:

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

Every official release should include a `SHA256SUMS.txt` file. Verify the SHA-256 hash of downloaded packages before use.

## Project status

- QH01: **ACCEPTED**
- Next development phase: **FD01**
- Current public channel: preview/binary release

## License

PersianWordAgent is **not open source**. Binary use is governed by the **PersianWordAgent Binary Distribution License 1.0** in `LICENSE.txt`.

Copyright © 2026 GOD13emad. All rights reserved.
