# 📦 Repository structure

This is a **source-first Excel UI library**. The Git diff, not an opaque
workbook, is the review artifact. This page lists what the repository holds and
where each kind of file belongs.

## File tree

```text
VBA-EXCEL_UI/
├─ .github/
├─ demo/
│  ├─ M_DEMO_BUILDER.bas
│  ├─ M_DEMO_CLEANUP_TESTS.bas
│  └─ M_EXCEL_UI_DEMO.bas
├─ docs/
│  ├─ EXCEL_CERTIFICATION.md
│  ├─ GOVERNANCE.md
│  ├─ ISSUE_LABELS.md
│  ├─ REPOSITORY_STRUCTURE.md
│  ├─ RIBBON_SDI_BEHAVIOR.md
│  ├─ VBA_CALL_CHECKER.md
│  ├─ VBA_HOUSE_STYLE.md
│  └─ WORKBOOK_SYNC.md
├─ images/
├─ src/
│  ├─ M_EXCEL_UI.bas
│  ├─ M_EXCEL_UI_RUNTIME.bas
│  ├─ M_EXCEL_UI_SNAPSHOT.bas
│  └─ M_EXCEL_UI_TITLEBAR.bas
├─ test/
│  └─ M_EXCEL_UI_REGRESSION_TESTS.bas
├─ tests/
├─ tools/
├─ AGENTS.md
├─ CHANGELOG.md
├─ CODE_OF_CONDUCT.md
├─ CONTRIBUTING.md
├─ INSTALLATION.md
├─ LICENSE
├─ README.md
├─ RELEASING.md
├─ SECURITY.md
├─ VERSION
└─ pyproject.toml
```

`tools/` holds the Python static gate and its helpers, the release tool, and
the Windows workbook tools (synchronization, export and certification), `tests/` the unit tests for
those tools, and `images/` the documentation images.

The source repository intentionally contains no versioned demo `.xlsm`. Release
binaries are attached to tagged GitHub Releases.

## Directories

| Location | Purpose |
|---|---|
| `src/` | Authoritative production modules for UI, runtime, snapshots, and title-bar control |
| `test/` | Regression and release-certification source |
| `demo/` | Demo and reproducible host setup |
| `tools/` | Repository/static validation |
| `tests/` | Unit tests for the static validation in `tools/` |
| `docs/` | VBA house style, governance, this page, issue and label rules, workbook synchronization, Excel certification, the VBA call checker, and measured Ribbon SDI behaviour |
| `images/` | Documentation images, the only approved home for binary files |
| `.github/` | Workflows, rulesets, repository settings, issue and pull-request templates, the label catalogue and its scripts |

The current tree is authoritative if a listed optional directory is not
present.

## Binary files

Binary files belong under `images/`. Anywhere else, `tools/check_repo.py`
rejects a tracked file that Git treats as binary (its `text` attribute is unset
in `.gitattributes`) or whose content contains a NUL byte, unless
`tools/binary_allowlist.txt` lists it with the SHA-256 of its exact content. The
digest makes the approval specific to that content: replacing the file fails
the gate until the entry is updated in the same change.

[VBA_HOUSE_STYLE.md](VBA_HOUSE_STYLE.md) holds the rules for the source files
themselves.
