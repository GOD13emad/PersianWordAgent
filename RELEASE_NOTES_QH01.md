# Persian WordAgent — QH01 Accepted Preview

Tag: `v2.0.0-qh01-preview`

This preview release is based on the QH01-accepted `2.0.0-m3.3-dev41` engine baseline.

## What is included

- Windows x64 desktop application
- Windows x64 command-line engine used by the desktop application
- runtime profiles required by the application
- binary license
- SHA-256 checksums

## Validated QH01 behavior

- Persian/RTL document handling
- governance proposal profile detection
- controlled-copy document repair
- Word-native TOC refresh
- TOC RTL/right alignment and page-number freshness checks
- table/caption pagination integrity gates
- targeted Gantt page-break correction
- Microsoft Word final rendering validation
- final desktop GUI regression and visual acceptance

## Acceptance evidence summary

- Engine baseline: `2.0.0-m3.3-dev41`
- Final GUI acceptance: PASS
- Manual GUI checks: 8/8 YES + explicit PASS
- Accepted validation document: 14 pages in Microsoft Word
- Stage 8 / Gantt caption / Gantt table: page 9
- Final explicit/cached Gantt page breaks: 0
- TOC page numbers: fresh
- DevRoot accepted graph: 47/47

## Important limitations

This is a preview release and not the final REL01 build.

The current accepted QH01 behavior is profile-driven. General template-authority mode is planned for the FD01 phase and is not claimed as complete in this preview.

## Requirements

- Windows x64
- Microsoft Word desktop for Word-native operations

## Distribution and license

This repository distributes official binaries only. Source code is not included.

Use of the binaries is governed by `PersianWordAgent Binary Distribution License 1.0` in `LICENSE.txt`.

Official downloads should be obtained only from the GitHub Releases page for this repository.
