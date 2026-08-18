# Changelog

All notable public binary-release changes to PersianWordAgent will be documented here.

## [v2.0.0-qh01-preview] - Preview

### Status

- QH01 explicitly accepted after final GUI and Microsoft Word validation.
- Engine baseline: `2.0.0-m3.3-dev41`.
- Public distribution model: binary release only.

### Validated capabilities

- Persian/RTL document processing.
- Profile-based governance proposal detection and controlled repair.
- Controlled-copy workflow with original-source protection.
- Word-native Table of Contents refresh.
- TOC RTL/right alignment and page-number freshness validation.
- Protected table pagination and caption integrity checks.
- Targeted Gantt page-break repair.
- Microsoft Word final rendering validation.
- Final GUI regression and visual acceptance.

### Validation summary

- Final Microsoft Word page count: 14 pages for the QH01 acceptance document.
- Stage 8, Gantt caption, and Gantt table validated on page 9.
- Final explicit/cached Gantt page breaks: 0.
- Final GUI test: 8/8 manual checks passed with explicit PASS verdict.
- DevRoot baseline: dev41, graph 47/47.
- Installed/Frozen states were not changed by QH01 acceptance.

### Distribution

- Source code is not included in the public repository or release package.
- Official binaries should be downloaded only from GitHub Releases.
- SHA-256 checksums are provided with official release assets.
