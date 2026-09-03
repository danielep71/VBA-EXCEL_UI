<div align="center">

# 📦 Installation and Upgrade Guide

### Install, validate, upgrade, recover, and remove the Windows Excel UI controller

[![Deployment](https://img.shields.io/badge/Deployment-Source--first-0969da?style=flat-square)](#deployment-model)
[![Validation](https://img.shields.io/badge/Validation-Required-d97706?style=flat-square)](#validation)
[![Security](https://img.shields.io/badge/Security-Review_before_enabling-d73a49?style=flat-square)](SECURITY.md)
[![Version](https://img.shields.io/badge/Version-VERSION_file-6f42c1?style=flat-square)](VERSION)
[![License](https://img.shields.io/badge/License-MIT-217346?style=flat-square)](LICENSE)

<br>

**Back up · Import one coherent version · Compile · Validate · Preserve caller state**

</div>

---

This guide covers installation, validation, upgrade, recovery, and removal of
**VBA Excel UI**.

> [!IMPORTANT]
> VBA source can execute with the user's Office permissions. Review the exact
> source or use a trusted tagged release, follow the organization's macro
> security policy, and never enable macros in an untrusted workbook.

---

## 🧭 Support baseline

| Item | Requirement |
|---|---|
| Host | Desktop Microsoft Excel for Windows |
| Office bitness | 32-bit and 64-bit Office |
| Version identity | Root `VERSION` file and the selected tag/commit |
| Source policy | Exported repository source is authoritative |
| Licence | MIT |
| Current deployment status | Four-module source-first library; all production modules must come from one release. |

Compatibility claims apply only to environments actually certified for the
selected release. Read [README.md](README.md), [CHANGELOG.md](CHANGELOG.md), and
the release notes before installation.

### Current release boundaries

Known defects and their correction state on this branch. A row marked
*corrected on branch* is not certified: certification requires a real
Windows Excel host and is tracked on the linked issue.

| Area | v1.1.2 boundary | v1.1.3 branch state |
|---|---|---|
| Ribbon restore | A changed active window can make restore target the wrong window instead of failing closed. | Open: [#23](https://github.com/danielep71/VBA-EXCEL_UI/issues/23) |
| Title-bar identity | v1.1.3 pairs the retained Excel Window with the hWnd read from that same object and fails closed when the pair no longer matches. | Corrected on branch; certification pending: [#45](https://github.com/danielep71/VBA-EXCEL_UI/issues/45) |
| Recycled hWnd | Registry slots retain Window generation identity, so equal style bits cannot authenticate a recycled handle. | Corrected on branch; certification pending: [#32](https://github.com/danielep71/VBA-EXCEL_UI/issues/32) |
| Captionless baseline | Show rejects zero and non-zero baselines without WS_CAPTION and confirms the live result by readback. | Corrected on branch; certification pending: [#6](https://github.com/danielep71/VBA-EXCEL_UI/issues/6) |
| Self-test ownership | The tagged self-test can clear a snapshot it has just refused because the caller owns it. | Corrected on the release branch: [#43](https://github.com/danielep71/VBA-EXCEL_UI/issues/43) |
| Quiet-update ownership | A suppressed or ignored write can be recorded as an achieved transition. | Corrected on the release branch: [#26](https://github.com/danielep71/VBA-EXCEL_UI/issues/26) |

<a id="deployment-model"></a>

## 🎯 Deployment model

The public facade is not a standalone installation. Static checks cannot certify Excel behavior; release certification requires a real Windows Excel host.

Choose one supported model and keep its source identity explicit:

| Model | Use when | Trust boundary |
|---|---|---|
| Embedded source | The component must travel with a workbook or add-in | Destination project contains the reviewed source |
| Tagged source | You build or integrate the component yourself | Tag/commit and exported files define identity |
| Published binary | The project explicitly ships a workbook/add-in asset | Hash, tag binding, and package smoke evidence are required |
| Development source | Focused testing or contribution work | Not a supported release unless the project says otherwise |

Do not combine files from different tags, commits, release assets, local exports,
or copied workbooks.

---

## 📂 Production source package

| Order | Repository source | VBE component | Responsibility |
|---:|---|---|---|
| 1 | `src/M_EXCEL_UI_RUNTIME.bas` | `M_EXCEL_UI_RUNTIME` | Host operations, diagnostics, result buffers, and quiet-update scope |
| 2 | `src/M_EXCEL_UI_TITLEBAR.bas` | `M_EXCEL_UI_TITLEBAR` | WinAPI declarations, owned title-bar bits, and frame refresh |
| 3 | `src/M_EXCEL_UI_SNAPSHOT.bas` | `M_EXCEL_UI_SNAPSHOT` | Snapshot state, capture/restore, and window identity |
| 4 | `src/M_EXCEL_UI.bas` | `M_EXCEL_UI` | Public `UI_*` facade and targeting orchestration |

Optional material is not part of the normal runtime unless stated otherwise:

| Source | Purpose |
|---|---|
| `test/M_EXCEL_UI_REGRESSION_TESTS.bas` | Regression and release certification |
| `demo/M_EXCEL_UI_DEMO.bas` | Demonstration actions |
| `demo/M_DEMO_BUILDER.bas` | Demo worksheet construction |

> [!CAUTION]
> A `.frm` and its `.frx` companion are one logical component. Keep them in
> the same directory during import, never import the `.frx` separately, and
> never process it as text.

---

## 🚀 Fresh installation

1. Back up the macro-enabled host and ensure Excel UI is currently recoverable.
2. Remove or replace any older complete production set; do not layer new modules over a mixed installation.
3. Import all four production modules in dependency-first order.
4. Compile the complete VBA project.
5. Run `UI_HideExcelUI` followed by `UI_ShowExcelUI`.
6. Capture, change, and restore a baseline with the public snapshot API.

### VBE import procedure

1. Open the destination workbook or add-in and press `Alt+F11`.
2. Select the intended project in Project Explorer.
3. Use **File → Import File…** for exported modules, classes, and forms.
4. Confirm component names match the repository source.
5. Run **Debug → Compile VBAProject**.
6. Save in a macro-capable format such as `.xlsm`, `.xlsb`, or `.xlam`
   when the project requires executable VBA.
7. Close and reopen the host before the clean-session smoke test.

Do not paste source into arbitrarily named modules when an exported component is
available. VBE attributes, component identity, form resources, and line endings
are part of a reproducible source installation.

---

<a id="validation"></a>

## ✅ Validation

A successful import is not sufficient evidence that the installation is correct.

- Run `Test_EXCEL_UI_RunReleaseCertification` for production/release validation.
- Treat any incomplete, skipped, failed, or cleanup-failed unit as not passing.
- Verify multi-window identity, closed/recreated windows, snapshot reset, and emergency recovery.
- Exercise title-bar behavior on the affected Office bitness and confirm unrelated style bits remain intact.
- Confirm `ScreenUpdating` and other caller-owned state are restored after success and failure.

### Minimum installation evidence

~~~text
Source tag or full commit SHA:
VERSION:
Files imported:
Excel version/build:
Office bitness:
Operating system:
Compile:
Consumer smoke:
Regression/certification:
Cleanup:
Skipped or unverified:
~~~

Treat a skipped, incomplete, cleanup-failed, or wrong-environment run as not
certified. Static checks and source review do not replace execution in Excel.

---

## ⬆️ Upgrade

Before upgrading:

1. read the complete version-to-version changelog;
2. back up the host and export any local modifications;
3. stop or clean up active component state;
4. identify every required production component;
5. decide whether stored configuration or generated assets are compatible.

- Replace all four production modules together from the same tag.
- Re-import the matching optional regression module when validating an upgrade.
- Review newly observable diagnostics and recovery behavior even when public signatures are unchanged.
- Never restore per-window state by collection index or reuse a stale snapshot as current authority.

After replacement, compile and repeat the full installation validation. Do not
claim an upgrade is non-breaking solely because VBA signatures compile.

### Local modifications

A locally modified copy is a fork. Diff it against the old and new exported
source, reapply changes deliberately, and retest. Do not overwrite it and assume
the local behavior survived.

---

## 🧯 Troubleshooting

| Symptom | Check |
|---|---|
| Compile error or missing procedure | Confirm every required component was imported from one version and optional dependencies are present. |
| Ambiguous name | Remove duplicate/legacy modules; do not paste new source beside old components. |
| Form missing controls or corrupt UI | Re-import the `.frm` with its exact adjacent `.frx`. |
| Behavior differs by workbook | Check caller, active-object, settings namespace, references, locale, and date-system assumptions. |
| 32/64-bit failure | Confirm the tested Office bitness and conditional WinAPI declarations. |
| Excel left altered after failure | Run the documented recovery/cleanup path; do not blindly force global state. |
| Security warning | Verify source origin, signature/hash where provided, trusted location policy, and macro settings. |
| Output differs from reference | Confirm exact version, inputs, parameterization, tolerance, environment, and reference independence. |

If recovery is uncertain, save user data separately, close Excel, reopen a clean
session, and reproduce with a minimal sanitized workbook before changing code.

Report suspected vulnerabilities privately under [SECURITY.md](SECURITY.md).

---

## 🗑️ Removal

1. Call `UI_ShowExcelUI` and clear any retained snapshot.
2. Remove all four production modules and optional demo/test modules.
3. Compile the remaining project and reopen Excel to confirm the normal shell is visible.

Removing files does not automatically remove workbook formulas, Ribbon XML,
registry settings, trusted-location configuration, cached add-ins, shortcuts,
scheduled callbacks, or other integrations. Remove only state the component
owns and document anything intentionally retained.

---

## 🔐 Security and privacy

- Obtain source and assets from the official repository or a verified release.
- Compare the selected tag, `VERSION`, release notes, and any published hash.
- Review VBA before enabling macros.
- Do not test with client, personal, regulated, or confidential workbooks.
- Inspect example and release workbooks for links, connections, names,
  properties, hidden content, and embedded code.
- Follow organizational macro, add-in, trusted-location, and signing policy.
- Report vulnerabilities through [SECURITY.md](SECURITY.md), not publicly.

---

## 📚 Related documentation

- [README.md](README.md) — capabilities, requirements, and public API
- [CHANGELOG.md](CHANGELOG.md) — version history and compatibility
- [CONTRIBUTING.md](CONTRIBUTING.md) — source and validation standards
- [RELEASING.md](RELEASING.md) — maintainer release and provenance procedure
- [SECURITY.md](SECURITY.md) — private vulnerability reporting
- [LICENSE](LICENSE) — MIT licence terms

---

### Installation principle

> Install one identifiable source version, compile it, exercise its real host
> behavior, and keep evidence of what was—and was not—validated.
