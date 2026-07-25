# 🔒 Security Policy

<p align="left">
  <img alt="Reporting" src="https://img.shields.io/badge/Reporting-Private-d97706">
  <img alt="Scope" src="https://img.shields.io/badge/Scope-VBA_WinAPI_and_demo_artifacts-6f42c1">
  <img alt="Stable release" src="https://img.shields.io/badge/Supported-Latest_tagged_release-217346">
  <img alt="Development" src="https://img.shields.io/badge/Development-Best_effort-lightgrey">
</p>

**VBA-EXCEL_UI** distributes plain-text VBA modules and a macro-enabled
demonstration workbook.

There is:

- no installer;
- no background service;
- no package manager;
- no third-party DLL shipped by the project;
- no credential store;
- no network service;
- no automatic update mechanism.

The attack surface is limited, but it is not zero. The project uses:

- VBA macros;
- fixed Excel 4 macro commands for Ribbon control;
- WinAPI calls for Excel main-window style management;
- a binary `.xlsm` demo artifact;
- process-wide Excel UI settings.

Responsible disclosure therefore matters.

> [!IMPORTANT]
> Hiding Excel UI elements is a presentation and workflow feature. It is **not**
> a security boundary, access-control mechanism, or substitute for workbook,
> VBA-project, operating-system, or organizational security controls.

---

## 📦 Supported versions

| Version or branch | Support status |
|---|---|
| Latest tagged release | ✅ Supported |
| Current release branch before publication | ⚠️ Release-candidate testing only |
| `main` | ⚠️ Best-effort development support |
| Older tags | ❌ Normally unsupported unless the issue also affects the latest release |
| Modified third-party copies | ❌ Unsupported |

Security fixes are normally prepared on a controlled branch and included in a
new tagged release.

When reporting an issue, identify the exact source state:

- a release tag; or
- the full Git commit SHA.

Do not report only “latest,” because repository branches can change after the
issue is observed.

---

## 📣 Reporting a vulnerability

**Do not open a public GitHub issue for a suspected security vulnerability.**

Use one of these private channels:

### 1. GitHub private vulnerability reporting

Where enabled:

1. Open the repository’s **Security** tab.
2. Select **Report a vulnerability**.
3. Submit the report privately.

### 2. Email the maintainer

```text
danielep71@gmail.com
```

Use a clear subject such as:

```text
Private security report — VBA-EXCEL_UI
```

Include:

- affected release tag or full commit SHA;
- affected file, module, and procedure;
- Excel product and version;
- Office 32-bit or 64-bit;
- Windows version;
- workbook and Excel window state;
- whether the official demo workbook was used;
- minimal reproduction steps;
- observed behavior;
- expected behavior;
- practical confidentiality, integrity, or availability impact;
- whether exploitation requires a modified workbook or untrusted macro source;
- any proposed mitigation;
- whether public disclosure has already occurred.

Do not attach workbooks containing:

- confidential information;
- client data;
- credentials;
- personal data;
- proprietary VBA;
- production connections or external links.

Provide a sanitized minimal reproduction where possible.

---

## 🎯 What qualifies as a security issue

Examples that should be reported privately include:

### Code execution and trust boundary

- execution of unintended code caused by repository-supplied source or artifacts;
- a repository-supplied workbook containing unexpected macros, links, connections,
  or embedded content;
- unsafe dynamic construction of macro commands;
- introduction of arbitrary `Shell` execution or external executable invocation;
- a path that allows untrusted input to select or construct WinAPI calls.

### Integrity

- unintended modification of workbook data, VBA projects, files, or external
  resources;
- title-bar or window-style changes that corrupt the Excel host state beyond the
  documented UI effect;
- failure behavior that persistently damages the Excel user interface after
  reasonable recovery attempts;
- a crafted call that applies UI changes outside the documented Excel process or
  scope.

### Confidentiality

- disclosure of workbook, environment, file, or user information beyond the
  documented result or diagnostic behavior;
- diagnostics that expose sensitive data unexpectedly;
- repository artifacts containing undisclosed personal, confidential, or
  machine-specific information.

### Availability

- a crafted input or state that causes persistent Excel hangs, uncontrolled loops,
  runaway resource consumption, or repeated crashes;
- a recovery failure that prevents practical restoration of the Excel interface
  and requires process termination;
- a title-bar or Ribbon path that creates an exploitable denial of service beyond
  an ordinary reproducible defect.

### Supply-chain integrity

- tampering with release artifacts;
- mismatch between a tagged release and the documented source;
- malicious content in a committed `.xlsm` artifact;
- compromised links or instructions that direct users to untrusted downloads.

When uncertain, report privately. The maintainer can reclassify the report safely.

---

## 🐞 Ordinary bugs

A serious defect is not automatically a security vulnerability.

Use a public issue for problems such as:

- an element does not show or hide correctly;
- a no-op write occurs unnecessarily;
- the Ribbon state cannot be read on one Excel version;
- the title bar renders incorrectly but does not create a concrete security or
  availability impact;
- the wrong failure count or message is returned;
- a snapshot restores the wrong per-window state without a security consequence;
- `ScreenUpdating` is restored incorrectly but Excel remains recoverable;
- documentation is inaccurate;
- a demo control is misaligned;
- performance is suboptimal but bounded.

Public bug reports should still avoid confidential workbooks and personal data.

---

## 🧭 Scope

### In scope

- `src/M_EXCEL_UI.bas`;
- `demo/M_EXCEL_UI_DEMO.bas`;
- `demo/M_DEMO_BUILDER.bas`;
- `test/M_EXCEL_UI_REGRESSION_TESTS.bas`;
- the official `demo/EXCEL_UI_DEMO.xlsm` artifact;
- repository release archives and attached release artifacts;
- fixed Excel 4 macro use for Ribbon management;
- WinAPI declarations and window-style handling;
- documented recovery and safe-use instructions;
- behavior that violates the project’s stated integrity boundary.

### Out of scope

- vulnerabilities in Microsoft Excel, Office, Windows, GitHub, or the VBA runtime;
- organization-controlled macro-security configuration;
- unrelated macros or add-ins in the host Excel process;
- malicious workbooks not supplied by this repository;
- user modifications to the source;
- copies downloaded from unofficial mirrors;
- unsupported historical snapshots;
- social-engineering attacks unrelated to project content;
- UI hiding treated as an access-control bypass;
- ordinary bugs without a concrete security impact.

A vulnerability in Excel or Windows should be reported to the relevant vendor.

---

## 🪟 Security considerations for UI and WinAPI behavior

### UI hiding is not protection

This library can hide interface elements, but it does not prevent:

- keyboard shortcuts;
- other macros;
- add-ins;
- the VBA Editor;
- workbook file manipulation;
- programmatic access through Excel automation;
- a knowledgeable user from restoring interface elements.

Do not use this project to enforce authorization or segregation of duties.

### Process-wide effects

Some managed properties are application-level and affect the current Excel
process, not only one workbook.

A host workbook should:

- document this effect;
- apply constrained UI only when appropriate;
- provide a recovery path;
- restore or show the managed UI before shutdown where practical.

### WinAPI title-bar handling

Window-style changes are sensitive because incorrect masks or handles can affect
the Excel frame.

Security-sensitive review is required for changes to:

- `GetWindowLong` / `GetWindowLongPtr`;
- `SetWindowLong` / `SetWindowLongPtr`;
- `SetWindowPos`;
- style masks;
- `Application.Hwnd` tracking;
- error handling around valid zero returns;
- frame refresh flags.

### Ribbon macro command

The Ribbon path uses a fixed Excel 4 macro command.

Do not change it to accept arbitrary user-controlled macro text. Any additional
Excel 4 macro use requires explicit security review and documentation.

---

## ⏱️ Disclosure process

This is a solo-maintained open-source project, so response times are best effort.

The expected process is:

1. The report is acknowledged.
2. The affected source and environment are identified.
3. The issue is reproduced where possible.
4. Security impact and affected versions are assessed.
5. A remediation and release strategy is agreed.
6. A fix is developed on a private or controlled branch when necessary.
7. Regression and recovery tests are added.
8. A corrected tagged release is published.
9. Public disclosure follows after users have had reasonable time to update.

Please allow reasonable time for investigation and remediation before public
disclosure.

Credit will be included in release notes when requested, unless anonymity is
preferred.

---

## 🧰 Safe-use guidance

### Obtain and inspect source safely

- Obtain source only from the official repository or a tagged release.
- Record the release tag or commit SHA used.
- Review plain-text `.bas` files before importing them.
- Treat macro-enabled workbooks as executable content.
- Review the official `.xlsm` demo before enabling macros.
- Do not enable macros in a workbook obtained from an untrusted mirror.

### Preserve macro security

- Keep Excel macro security at the organization’s approved level.
- Do not lower macro-security settings solely to use this project.
- Prefer trusted locations, signed macros, or organizational deployment controls
  where required by policy.
- Do not instruct users to disable Protected View globally.
- Do not bypass Mark-of-the-Web controls without verifying file provenance.

### Compile and test

After importing:

```text
VBA Editor → Debug → Compile VBAProject
```

Then run:

```vb
Test_EXCEL_UI_RunCore
Test_EXCEL_UI_RunTitleBarOnly
Test_EXCEL_UI_RunAll
```

Perform emergency recovery validation:

```vb
UI_HideExcelUI
UI_ShowExcelUI
```

### Maintain a recovery path

Host solutions should keep an accessible recovery procedure:

```vb
Public Sub RestoreManagedExcelUI()
    UI_ShowExcelUI
End Sub
```

Where appropriate, expose it through:

- the VBA Editor;
- a trusted custom Ribbon control;
- the Quick Access Toolbar;
- a documented startup or shutdown recovery path.

### Protect sensitive information

- Do not include client or personal data in issues.
- Sanitize screenshots.
- Remove external links and connections from demonstration workbooks.
- Do not embed credentials in VBA.
- Do not log sensitive workbook values in UI failure messages.

### Windows downloaded-file blocking

If Windows marks downloaded files as blocked:

1. verify the official source and release;
2. inspect the files;
3. follow organizational policy;
4. use **Properties → Unblock** only after provenance has been established.

---

## 🔍 Verifying a release

For controlled use:

1. download from the official repository release page;
2. record the tag;
3. inspect the release notes;
4. review the `.bas` modules;
5. compare repository content with the tag;
6. scan the `.xlsm` artifact under organizational policy;
7. compile the VBA project;
8. run the regression harness;
9. document the tested Excel and Windows environment.

Checksums may be added to future releases. Until then, the Git tag and repository
history are the primary source-integrity references.

---

## 📚 Related policies

- [Contributing Guidelines](CONTRIBUTING.md)
- [Code of Conduct](CODE_OF_CONDUCT.md)
- [MIT License](LICENSE)
- [Project README](README.md)

---

## 👤 Maintainer

Maintained by **Daniele Penza**.
