<div align="center">

# 🔒 Security Policy

### Trust boundaries, responsible disclosure, safe deployment, and release integrity for VBA-EXCEL_UI

[![Reporting](https://img.shields.io/badge/reporting-private-d97706?style=for-the-badge)](#-reporting-a-vulnerability)
[![Support](https://img.shields.io/badge/support-latest_tagged_release-217346?style=for-the-badge)](#-supported-versions)
[![Platform](https://img.shields.io/badge/platform-Excel_VBA_%2F_Windows-0078D6?style=for-the-badge)](#-security-model)
[![Scope](https://img.shields.io/badge/scope-runtime_release_and_automation-6f42c1?style=for-the-badge)](#-scope)
[![Automation](https://img.shields.io/badge/automation-least_privilege-d73a49?style=for-the-badge)](#-repository-automation-and-runner-security)

<br>

**Source-first trust · Fixed Ribbon commands · Owned-bit WinAPI mutation · Explicit snapshot ownership · Private disclosure · Exact-source evidence**

<br>

[Supported versions](#-supported-versions)
&nbsp;·&nbsp;
[Report privately](#-reporting-a-vulnerability)
&nbsp;·&nbsp;
[Security scope](#-scope)
&nbsp;·&nbsp;
[Runtime boundaries](#-runtime-security-boundaries)
&nbsp;·&nbsp;
[Supply chain](#-supply-chain-and-release-integrity)
&nbsp;·&nbsp;
[Automation](#-repository-automation-and-runner-security)
&nbsp;·&nbsp;
[Verify a release](#-verifying-a-release)

</div>

---

**VBA-EXCEL_UI** distributes reviewable VBA source for controlling selected
Microsoft Excel interface elements on Windows. A macro-enabled demonstration
workbook may also be distributed as a GitHub Release asset.

The production component runs with the privileges already granted to Microsoft
Excel and the current Windows user.

There is no:

- privileged installer;
- background service;
- automatic updater;
- production network client;
- production shell execution;
- bundled third-party DLL;
- credential store;
- elevated Windows service or driver;
- package manager required by the VBA runtime;
- machine-level registration performed by the component.

The component does:

- execute fixed Excel 4 macro commands for Ribbon visibility;
- read and write selected Excel Application and Window properties;
- call user32 and kernel32 for native frame-style management;
- retain in-memory Excel Window references while a snapshot exists;
- temporarily suppress Application.ScreenUpdating during UI update passes;
- change UI state that can affect more than the calling workbook;
- ship source that users import into trusted macro-enabled workbooks or add-ins.

The attack surface is therefore focused, but it is not zero.

> [!IMPORTANT]
> VBA-EXCEL_UI is **not a security boundary**.
>
> Hiding the Ribbon, Formula Bar, Status Bar, headings, workbook tabs,
> gridlines, scroll bars, or title bar does not enforce authorization, protect
> worksheet data, restrict the VBA Editor, prevent automation, or isolate one
> workbook from other trusted VBA in the same Excel process.
>
> The component’s diagnostics, identity checks, owned-bit frame writes,
> snapshot lifecycle, and recovery paths are operational safety controls. They
> are not a sandbox against malicious VBA already running with the same Excel
> privileges.

---

## 🧭 Security model

The project assumes:

~~~text
Microsoft Excel and Windows are trusted
the VBA project containing the component is trusted
the current Windows user is authorized to run that VBA project
macros are enabled through an approved trust mechanism
callers invoke the public API intentionally
visibility arguments and target scopes come from trusted host code
~~~

The project does **not** assume:

~~~text
hidden Excel UI is inaccessible
every workbook open in the Excel process is cooperative
every Excel property is local to one workbook or one window
Application.Hwnd identifies the same window after activation changes
a non-zero hWnd proves that the handle still belongs to the same Excel Window
IsWindow proves object identity
an Excel Window object and a native hWnd are interchangeable identities
a successful native style write proves the requested frame is visibly correct
a VBA project reset restores native frame styles or releases Excel process state
a passing static workflow proves Excel/VBA runtime behavior
a release workbook was built from tagged source merely because both are named
  in the same release
~~~

These boundaries explain the design:

- application-level and window-level settings are treated separately;
- window-level changes use explicit target scopes;
- snapshot restoration retains Excel Window objects rather than relying on
  collection indexes;
- title-bar mutation merges only the five style bits owned by the component;
- native zero returns are interpreted using the relevant API contract and
  GetLastError;
- frame changes are refreshed through SetWindowPos without moving, resizing, or
  changing Z-order;
- structured calls expose failure counts and ordered detail rather than
  presenting success unconditionally;
- UI_ShowExcelUI remains the independent show-all recovery path;
- release assurance distinguishes static source checks, Excel execution
  evidence, artifact identity, and exact-source provenance.

---

## 📦 Supported versions

Security fixes are normally applied to the **latest tagged release**.

| Source state | Security support |
|---|---|
| **Latest tagged release** | ✅ Supported |
| Release branch before publication | ⚠️ Release-candidate testing / best effort |
| main | ⚠️ Development branch / best effort |
| Older tagged releases | ❌ Normally unsupported; upgrade first |
| Modified forks or copied modules | ❌ Unsupported unless the issue reproduces in official source |
| Unofficial binary mirrors | ❌ Unsupported |

When reporting, identify the exact affected source with one of:

~~~text
release tag
full 40-character commit SHA
~~~

Do not report only:

~~~text
latest
current
main from yesterday
~~~

Those descriptions can identify a different source state after the report is
submitted.

### Security-fix policy

A confirmed vulnerability may result in:

- a controlled fix branch;
- regression or fault-injection coverage;
- a private draft security advisory;
- a corrected tagged release;
- replacement or withdrawal of an affected release asset;
- mitigation guidance;
- changes to workflow or runner permissions;
- a coordinated public disclosure.

Older releases are not normally patched in place.

---

## 📣 Reporting a vulnerability

**Do not open a public GitHub issue for a suspected vulnerability.**

Do not publish:

~~~text
exploit code
weaponized macro-enabled workbooks
credentials or personal access tokens
private signing keys
client workbooks or client data
proprietary VBA
proof-of-concept runner escapes
sensitive workstation or network details
~~~

in a public issue, pull request, discussion, Wiki page, release thread, or
certification artifact.

### Option 1 — GitHub private vulnerability reporting

If private vulnerability reporting is enabled:

~~~text
Repository
→ Security
→ Report a vulnerability
~~~

Submit the report there.

### Option 2 — email the maintainer

~~~text
danielep71@gmail.com
~~~

Suggested subject:

~~~text
Private security report — VBA-EXCEL_UI
~~~

### Include, where relevant

- affected release tag or full commit SHA;
- exact file, module, procedure, or workflow;
- Excel version and build;
- Office 32-bit or 64-bit;
- Windows version;
- workbook and Excel-window configuration;
- number of workbooks and windows involved;
- active window at capture, mutation, and restore;
- relevant target scope;
- whether a UI snapshot already existed;
- whether a VBA project reset or unhandled error occurred;
- whether the official demo asset was used;
- whether the report concerns source, a release asset, CI, or a runner;
- minimal reproduction steps;
- observed behavior;
- expected behavior;
- concrete confidentiality, integrity, availability, credential, or
  supply-chain impact;
- whether exploitation requires already-trusted malicious VBA;
- any proposed mitigation;
- whether public disclosure has already occurred.

### Safe reproduction material

Prefer:

~~~text
sanitized workbook
minimal plain-text VBA module
plain-text reproduction steps
structured FailureCount and FailureList output
sanitized certification text or JSON
screenshots with names and data removed
hashes and exact release metadata
~~~

Do not attach a production workbook merely because it reproduces the problem.

Remove:

- client or personal data;
- workbook connections and external links;
- credentials;
- proprietary VBA unrelated to the defect;
- file paths and workbook names that reveal sensitive information;
- hidden sheets or metadata not needed for reproduction.

---

## ⏱️ What to expect

This is a solo-maintained open-source project.

Response times are **best effort**, not a contractual service-level agreement.

The expected process is:

1. acknowledge the report;
2. identify the affected source, artifact, and environment;
3. reproduce where practical;
4. distinguish correctness and operational safety from security impact;
5. determine affected versions and artifacts;
6. agree on immediate mitigation where needed;
7. develop remediation on a controlled branch;
8. add regression or fault-injection coverage where appropriate;
9. validate on the relevant Excel and Windows host;
10. publish a corrected release or other remedy;
11. disclose publicly after users have had reasonable time to update.

Credit can be included in release notes or an advisory when requested. Anonymous
credit is also acceptable.

Please allow reasonable investigation and remediation time before public
disclosure.

---

## 🎯 What qualifies as a security issue

When uncertain, report privately. The maintainer can safely reclassify a report
as an ordinary defect.

### 1. Unexpected code execution

Examples:

- official source or an official release workbook executes code outside the
  documented component behavior;
- user-controlled text reaches ExecuteExcel4Macro instead of the component’s
  fixed Ribbon commands;
- a repository-supplied workbook contains undisclosed macros, external links,
  connections, embedded payloads, or executable content;
- a project change introduces Shell, process creation, arbitrary
  Application.Run dispatch, or external executable invocation without an
  explicit documented trust model;
- crafted input can select a native API target or call outside the documented
  fixed implementation.

### 2. Integrity

Examples:

- a crafted state causes the component to mutate a different native window
  outside its documented Excel target and creates a concrete security impact;
- frame-style logic overwrites Windows style bits the component does not own;
- recovery reports trustworthy success while leaving security- or
  integrity-sensitive Excel process state in an unknown condition;
- a non-owner workflow can destroy or replace another workflow’s snapshot in a
  way that crosses a real trust boundary;
- repository tooling or automation makes a false source, test, tag, or artifact
  provenance claim;
- an official release asset materially differs from what its release claims.

### 3. Confidentiality

Examples:

- diagnostics expose workbook data, sensitive workbook or window names,
  workstation paths, environment data, or personal information unexpectedly;
- certification output publishes client-specific information;
- a release workbook contains confidential, personal, or machine-specific
  information;
- a workflow artifact or log exposes a token, signing secret, private path, or
  runner detail;
- demo or test tooling uploads an unsanitized workbook.

### 4. Availability

Examples:

- crafted input causes persistent Excel hangs, repeated crashes, or uncontrolled
  loops without a practical recovery path;
- native frame mutation persistently leaves Excel unusable after documented
  recovery attempts;
- an unbalanced quiet-update or cleanup path persistently suppresses required
  Excel process behavior;
- a release or regression workbook repeatedly re-enters a constrained UI state
  and prevents practical recovery;
- a self-hosted Excel runner can be persistently compromised or left with
  orphaned Excel processes that contaminate later jobs.

### 5. Supply-chain integrity

Examples:

- tampered GitHub Release assets;
- unauthorized replacement of a published demo workbook;
- compromised repository workflow;
- compromised or unexpectedly changed third-party action;
- malicious redirect in download instructions;
- mismatch between release notes, source tag, certification evidence, and
  published artifacts;
- forged or misleading release provenance;
- compromise of a Windows/Excel runner used to certify releases.

### 6. Credentials and automation

Examples:

- repository or runner secrets exposed in logs or artifacts;
- workflow permissions broadened beyond what the job requires;
- untrusted pull-request code gaining access to a persistent self-hosted runner;
- signing keys becoming available to ordinary test jobs;
- a high-privilege token becoming available to repository scripts that do not
  require it.

---

## 🐞 Ordinary bugs

A serious correctness defect is not automatically a security vulnerability.

Use a public issue for ordinary defects such as:

- a UI element does not show or hide correctly;
- a no-op write occurs unnecessarily;
- Ribbon state cannot be read on one Excel version;
- Ribbon restoration targets the wrong active window without a concrete
  security impact;
- a title-bar identity or recycled-handle defect remains recoverable and has no
  concrete confidentiality, integrity, or availability impact;
- a structured call returns an inaccurate failure count or message;
- a snapshot restores the wrong ordinary window state without crossing a trust
  boundary;
- ScreenUpdating is restored incorrectly but Excel remains practically
  recoverable;
- documentation is inaccurate;
- the demo is incomplete or misaligned;
- performance is suboptimal but bounded;
- a regression or certification case is missing.

Known corrective work is tracked publicly, including:

- Ribbon wrong-target restoration
  ([#23](https://github.com/danielep71/VBA-EXCEL_UI/issues/23));
- title-bar Window and hWnd identity pairing
  ([#45](https://github.com/danielep71/VBA-EXCEL_UI/issues/45));
- recycled-hWnd registry collision
  ([#32](https://github.com/danielep71/VBA-EXCEL_UI/issues/32));
- captionless title-bar show false success
  ([#6](https://github.com/danielep71/VBA-EXCEL_UI/issues/6));
- certification self-test snapshot ownership
  ([#43](https://github.com/danielep71/VBA-EXCEL_UI/issues/43)).

If a defect can be intentionally driven across a concrete confidentiality,
integrity, availability, credential, or supply-chain boundary, report it
privately even when a related correctness issue is already public.

Public reports must still exclude confidential workbooks, credentials, personal
data, and unsafe proof-of-concept artifacts.

---

## 🛠️ Scope

### In scope — production source

~~~text
src/M_EXCEL_UI.bas
src/M_EXCEL_UI_RUNTIME.bas
src/M_EXCEL_UI_SNAPSHOT.bas
src/M_EXCEL_UI_TITLEBAR.bas
~~~

Future production modules introduced under src are also in scope.

### In scope — regression and demo source

~~~text
test/M_EXCEL_UI_REGRESSION_TESTS.bas
demo/M_EXCEL_UI_DEMO.bas
demo/M_DEMO_BUILDER.bas
~~~

A defect in demo or test code is security-relevant when it can:

~~~text
execute unexpected code
damage unrelated workbook or process state
leak sensitive data
misrepresent release-quality evidence
compromise a release runner
package undisclosed executable content
~~~

### In scope — repository tooling

~~~text
tools/check_repo.py
tools/reformat.py
tools/public_api_manifest.txt
~~~

Future committed release, build, regression, or provenance tooling is also in
scope.

### In scope — repository automation

Current workflow:

~~~text
.github/workflows/static-checks.yml
~~~

Future Windows/Excel execution workflows, self-hosted runner configuration,
release signing, and attestation automation are in scope.

### In scope — official release artifacts

Official GitHub Release assets are in scope, including demo workbooks named
using the documented pattern:

~~~text
EXCEL_UI_DEMO_v<major>.<minor>.<patch>.xlsm
~~~

Where published for a release, these are also in scope:

~~~text
checksums
release manifests
certification text or JSON
release notes
provenance and execution-certification claims
~~~

### In scope — runtime integrations

- fixed Excel 4 macro execution for Ribbon state;
- Excel Application and Window property access;
- user32 and kernel32 declarations;
- native frame-style reads, owned-bit merges, writes, and refresh;
- Window and hWnd identity pairing;
- snapshot capture, replacement, retention, restoration, and clearing;
- retained Excel Window object references;
- target-scope resolution under Excel’s SDI window model;
- Application.ScreenUpdating suppression and restoration;
- structured failure count and diagnostic-list integrity;
- emergency show-all recovery;
- runtime tests and fault-injection seams.

### Out of scope

- vulnerabilities in Microsoft Excel, Office, Windows, GitHub, Python, or VBA
  themselves;
- organization-controlled macro-security configuration;
- malicious VBA not supplied by this repository;
- unrelated workbooks or add-ins in the host Excel process;
- user modifications that do not reproduce in official source;
- unofficial mirrors or repackaged workbooks;
- old unsupported tags where the issue does not affect a supported version;
- social engineering unrelated to project content;
- treating hidden UI as an access-control bypass;
- a trusted caller deliberately invoking documented UI changes;
- ordinary correctness bugs without concrete security impact.

A vulnerability in Excel, Windows, GitHub, or another platform should be
reported to the relevant vendor.

---

## 🛡️ Runtime security boundaries

### 1. UI hiding is not protection

The component can hide selected interface elements. It does not prevent:

~~~text
keyboard shortcuts
other macros
Excel add-ins
the Visual Basic Editor
Office automation
direct workbook-file access
programmatic worksheet access
a knowledgeable user from restoring interface elements
~~~

Do not use VBA-EXCEL_UI to implement:

- authorization;
- authentication;
- segregation of duties;
- data-loss prevention;
- worksheet confidentiality;
- VBA-project protection;
- kiosk security;
- anti-tamper controls.

Use Excel, Windows, file, identity, and organizational security controls for
those purposes.

### 2. Same-process trust

The component has no privilege boundary separating it from other VBA running in
the same Excel process.

Trusted VBA can:

- call the public API;
- change the same Excel properties directly;
- activate a different workbook window;
- use ExecuteExcel4Macro independently;
- invoke the same Windows APIs;
- reset the VBA project;
- clear or replace project-level state.

Defensive checks reduce accidental misuse. They do not constrain malicious code
that already has equivalent Excel and Windows-user privileges.

### 3. Application, window, and active-window effects

Managed state does not have one universal scope.

| Surface | Effective scope |
|---|---|
| Status Bar | Current Excel Application instance |
| Scroll Bars | Current Excel Application instance |
| Formula Bar | Current Excel Application instance |
| Headings | Selected Excel Window targets |
| Workbook Tabs | Selected Excel Window targets |
| Gridlines | Selected Excel Window targets |
| Ribbon | Active-window command context |
| Title bar | Native top-level Excel frame selected by the implementation |

TargetScope controls headings, workbook tabs, and gridlines. It does not make
application-level state workbook-local and does not provide a native target
argument for Ribbon commands.

A host solution should:

- document cross-workbook effects;
- avoid constraining unrelated user sessions;
- coordinate active-window changes;
- provide recovery independent of snapshot state;
- restore or show managed UI before shutdown where practical.

### 4. Fixed Excel 4 Ribbon commands

Ribbon visibility uses Application.ExecuteExcel4Macro because the Excel object
model does not provide an equivalent supported hide/show property.

The production implementation constructs only these fixed command forms:

~~~text
Show.TOOLBAR("Ribbon",True)
Show.TOOLBAR("Ribbon",False)
~~~

and reads Ribbon state using a fixed Get.ToolBar expression.

The Boolean state is not inserted as arbitrary caller-controlled macro text.

Do not change this design to accept:

~~~text
worksheet-derived command text
file-derived command text
network-derived command text
arbitrary caller-supplied Excel 4 expressions
~~~

Any new Excel 4 macro use requires explicit security review and documentation.

The fixed-command design reduces command-injection risk. It does not make legacy
Excel 4 macro execution a general-purpose safe evaluation mechanism.

### 5. Native title-bar mutation

The title-bar module calls:

~~~text
GetWindowLong / GetWindowLongPtr
SetWindowLong / SetWindowLongPtr
SetWindowPos
IsWindow
GetLastError
SetLastError
~~~

The component claims only these frame-style bits:

~~~text
WS_CAPTION
WS_SYSMENU
WS_THICKFRAME
WS_MINIMIZEBOX
WS_MAXIMIZEBOX
~~~

Writes must preserve every unowned style bit. SetWindowPos refreshes the
non-client frame without moving, resizing, or changing Z-order.

Security-sensitive changes include:

- PtrSafe and 32-bit/64-bit declarations;
- Long, LongPtr, and native return handling;
- the owned-style mask;
- valid-zero and GetLastError handling;
- hWnd acquisition and lifetime;
- Window/hWnd pairing;
- handle-reuse defenses;
- style merge logic;
- baseline selection;
- frame-refresh flags;
- refresh-debt and retry behavior;
- error paths after a successful style write.

IsWindow proves that a handle currently refers to a window. It does not by
itself prove that the window is the same Excel Window originally captured.

A native declaration or identity defect is often a correctness or availability
issue rather than a privilege-escalation vulnerability. Reports should identify
the concrete impact.

### 6. Snapshot identity and ownership

The project maintains one in-memory snapshot slot per loaded VBA project.

The snapshot can retain:

- application-level visibility values;
- one retained Excel Window reference per captured window;
- per-window visibility values;
- Ribbon state;
- a title-bar Window reference, native hWnd, and frame state;
- flags identifying which reads were valid.

Ownership is a host-application responsibility:

~~~text
capture replaces the previous snapshot
restore retains the snapshot
clear releases retained Window references
project reset destroys the VBA snapshot state
new windows are not part of an earlier snapshot
closed or recreated windows can make restoration incomplete
~~~

Do not let a helper, test, or nested workflow capture or clear state owned by
another caller.

Window object identity, Window.hWnd, Application.Hwnd, and a native handle’s
liveness are related evidence. No single one should be treated as complete
identity proof in every SDI lifecycle.

### 7. Structured diagnostics

The WithResult APIs expose:

~~~text
Boolean success
FailureCount
optional ordered FailureList entries
~~~

FailureCount is authoritative. FailureList is best effort because allocating or
growing the list can fail under memory pressure.

Diagnostics may include:

- procedure or stage names;
- Excel error descriptions;
- workbook or window labels;
- native error numbers;
- host-version information in certification output.

Treat diagnostic and certification output as potentially sensitive operational
data. Sanitize it before publishing.

Structured diagnostics improve observability. They are not tamper-proof audit
records or cryptographic attestations.

### 8. Quiet-update ownership

UI update and snapshot-restore paths can temporarily set:

~~~text
Application.ScreenUpdating = False
~~~

The runtime records the caller-visible baseline and restores ScreenUpdating only
when it changed it.

ScreenUpdating is Excel-process state, not workbook-local state.

Changes to quiet-scope entry, cleanup, error propagation, or nested behavior
require review for:

- baseline preservation;
- cleanup on every exit;
- caller-owned False state;
- project-reset consequences;
- release-test cleanup evidence.

An interrupted VBA project can bypass normal cleanup. UI_ShowExcelUI and, where
needed, a complete Excel-process restart remain operational recovery options.

### 9. Test seams and regression behavior

The source includes internal fault-injection seams so regression tests can
exercise failures that Windows or Excel cannot produce deterministically.

These seams are test infrastructure, not authentication-protected capabilities.
Code already running in the same trusted VBA project normally has equivalent
privileges.

Keep test seams:

~~~text
non-public where practical
one-shot or explicitly resettable
deterministic
cleared on cleanup
unable to leak silently into later cases
documented as test infrastructure
~~~

The regression harness manipulates real Excel state. It can create windows,
change active windows, mutate title-bar styles, change Ribbon visibility,
suppress ScreenUpdating, inject failures, create result files, and retain or
clear snapshots.

Run it only in a controlled Excel session.

---

## ⚠️ Current correctness and assurance boundaries

The latest tagged release can contain known correctness limitations tracked in
the public issue backlog. These limitations do not automatically become
security vulnerabilities, but they must not be hidden behind stronger security
or certification claims.

For the current v1.1.2 baseline:

- Ribbon snapshot restore can act on the wrong active window
  ([#23](https://github.com/danielep71/VBA-EXCEL_UI/issues/23));
- title-bar restore requires stronger Window/hWnd pairing
  ([#45](https://github.com/danielep71/VBA-EXCEL_UI/issues/45));
- native-handle reuse can collide with same-style registry state
  ([#32](https://github.com/danielep71/VBA-EXCEL_UI/issues/32));
- a non-zero captionless baseline can produce false show success
  ([#6](https://github.com/danielep71/VBA-EXCEL_UI/issues/6));
- the certification self-test can destroy a caller-owned snapshot
  ([#43](https://github.com/danielep71/VBA-EXCEL_UI/issues/43)).

Release-assurance work also remains open for cleanup proof, mandatory case
inventory, exact-source evidence, complete public-API contract gating,
documentation closure, and exact-head review/certification.

Until the relevant corrections are released and certified:

- keep a trusted UI_ShowExcelUI recovery macro;
- do not run certification while another workflow owns the snapshot;
- test multi-window identity transitions in a disposable session;
- do not infer visible correctness solely from a True return;
- bind every release claim to the exact source actually tested;
- distinguish a public correctness limitation from a privately reportable
  exploit with concrete security impact.

Private assessments and audit material are not public security artifacts. Do
not commit or publish them without explicit authorization. Public issues,
release notes, advisories, and approved evidence must be self-contained.

---

## 📦 Macro-enabled release artifacts

The repository is source-first.

The authoritative implementation is the tagged exported source:

~~~text
src/M_EXCEL_UI.bas
src/M_EXCEL_UI_RUNTIME.bas
src/M_EXCEL_UI_SNAPSHOT.bas
src/M_EXCEL_UI_TITLEBAR.bas
~~~

A demo workbook is executable Office content.

> [!WARNING]
> A familiar filename is not proof of authenticity.
>
> Confirm the official release, inspect the asset, and compare any published
> digest before enabling macros.

The source tree intentionally does not track generated XLSM or XLAM binaries.
Official demo workbooks are distributed through GitHub Releases.

For controlled environments, import and compile the tagged plain-text source in
a clean workbook rather than relying solely on a prebuilt workbook.

The current public demo can lag behind the production source and may not
exercise the current API or corrective behavior. Read the release notes before
using a demo as evidence of the component’s behavior.

---

## 🔗 Supply-chain and release integrity

### Trusted distribution

Obtain source and release assets from:

- the official repository;
- the official GitHub Releases page;
- a specific official tag or full commit.

Do not rely on:

~~~text
third-party mirrors
files forwarded by email without provenance
renamed macro-enabled workbooks
unofficial package sites
blog attachments
binary copies whose hashes conflict with the official release
~~~

### Source-first review

For controlled deployment, review:

~~~text
the four production BAS modules
the exact tag or full commit SHA
CHANGELOG.md
INSTALLATION.md
SECURITY.md
the relevant workflow and repository-checking code
~~~

Review demo and regression source when those artifacts are part of the
deployment or evidence chain.

### What hashes and manifests prove

A SHA-256 digest establishes the identity of the file that was hashed.

It does not automatically establish:

~~~text
who built the file
which source was imported
whether the VBA project compiled
whether Excel executed the tests
whether the workbook contains only the documented modules
whether the build was reproducible
~~~

A release manifest can bind declared metadata such as:

~~~text
tag and commit SHA
source hashes
release-asset hashes
Excel build
Office bitness
certification result
~~~

but only to the extent that the release process generated and verified those
fields correctly.

Distinguish:

~~~text
tagged source identity
static-check evidence
Excel execution evidence
workbook file identity
source-to-workbook build provenance
review and approval evidence
~~~

These are separate claims.

### Current provenance boundary

The repository does not currently provide a complete reproducible,
cryptographically attested source-to-XLSM build chain.

Where a release publishes no checksum or manifest, do not invent one or imply
that GitHub hosting alone binds the binary to source.

Exact-source certification and exact-head release evidence are active
hardening areas. A release must not certify one tree and tag or publish another.

### Signing and attestations

Do not describe a release as signed, attested, or reproducibly built unless the
corresponding control is actually deployed and verifiable.

Possible future improvements include:

~~~text
signed annotated tags
release attestations
VBA project or workbook signing
mandatory SHA-256 publication
release manifests
immutable release-asset policy
stronger source-to-artifact provenance
~~~

---

## 🔍 Verifying a release

For controlled use:

1. obtain source and assets from the official repository;
2. record the release tag and full commit SHA;
3. inspect the release notes and CHANGELOG.md;
4. inspect the tagged production source;
5. treat every XLSM or XLAM as executable Office content;
6. inspect the official asset under organizational policy;
7. when a SHA-256 is published, compute the local digest:

   ~~~text
   certutil -hashfile "EXCEL_UI_DEMO_vX.Y.Z.xlsm" SHA256
   ~~~

8. compare it character for character with the official release value;
9. inspect any published manifest and confirm its tag, SHA, environment, and
   asset fields describe the intended release;
10. compile the exact VBA source:

    ~~~text
    VBA Editor → Debug → Compile VBAProject
    ~~~

11. run the required certification runner in a controlled Excel instance:

    ~~~vb
    Test_EXCEL_UI_RunReleaseCertification
    ~~~

12. preserve the runner-generated text and JSON evidence;
13. record Excel version/build, Office bitness, Windows version, window
    configuration, cases, failures, skipped cases, cleanup result, and exact
    source SHA;
14. require every mandatory case and cleanup condition to pass;
15. verify the static workflow or local static gate for the same exact SHA;
16. smoke-test the actual release workbook when that workbook is the deployed
    artifact;
17. invalidate and repeat evidence after any source or release-relevant
    documentation change.

A hash mismatch means the file is not the file represented by the published
digest. Treat that as a supply-chain concern.

A passing source test does not automatically certify workbook packaging.

---

## 🧰 Safe-use guidance

### 1. Preserve Excel macro security

- keep Excel macro security at the organization-approved level;
- do not disable Protected View globally;
- do not weaken Trust Center settings solely to use this component;
- use Trusted Locations or signed VBA only where policy supports them;
- unblock downloaded files only after provenance and policy checks.

### 2. Import one exact production package

Import all four production modules from the same tag or commit.

Do not mix:

~~~text
one module from main
another module from a tag
demo source from a different version
copied snippets with exported modules
~~~

Compile the complete destination VBA project after import.

### 3. Treat target scope as an integrity decision

Use the narrowest UIWindowTargetScope that matches the host workflow.

Document when a command intentionally affects:

- the active Excel window;
- every window of the active workbook;
- every current Excel window;
- application-level Excel properties.

TargetScope is a behavioral scope, not an access-control boundary.

### 4. Keep independent recovery

Keep a trusted recovery macro:

~~~vb
Public Sub RestoreManagedExcelUI()
    UI_ShowExcelUI
End Sub
~~~

Make it accessible through a controlled route such as:

- the Visual Basic Editor;
- a trusted custom Ribbon control;
- the Quick Access Toolbar;
- documented startup or shutdown recovery.

Do not rely exclusively on snapshot restore.

### 5. Coordinate snapshot ownership

- capture only when the snapshot slot is available;
- do not overwrite another workflow’s baseline;
- inspect structured capture and restore results;
- clear the snapshot when its owner has finished;
- release retained Window references before workbook or add-in shutdown;
- use UI_ShowExcelUI after uncertain state or project reset.

### 6. Test native and multi-window behavior safely

- save work before running title-bar or release tests;
- close unrelated workbooks;
- use a disposable workbook or Excel instance;
- test window activation, closure, recreation, and recovery;
- confirm ScreenUpdating and snapshot cleanup;
- restart Excel after uncertain native or process state.

### 7. Protect sensitive information

- do not include client or personal data in issues;
- sanitize screenshots;
- remove external links and connections from demo workbooks;
- do not embed credentials in VBA;
- do not log sensitive worksheet values;
- sanitize workbook/window names and local paths in diagnostics;
- review certification JSON and text before publication.

### 8. Handle downloaded-file blocking safely

If Windows marks a downloaded file as blocked:

1. confirm it came from the official release;
2. inspect its source and available digest;
3. follow organizational policy;
4. unblock only after provenance has been established.

Do not tell users to disable Mark-of-the-Web or Protected View globally.

---

## 🔑 Repository automation and runner security

The repository currently contains one software-quality workflow:

~~~text
.github/workflows/static-checks.yml
~~~

It:

- runs on a GitHub-hosted Ubuntu runner;
- checks exported source as text;
- uses the repository’s Python tooling;
- installs no project package dependencies;
- declares contents: read;
- does not require repository secrets;
- does not execute Excel or VBA;
- does not package or publish a release workbook.

This is intentionally a low-privilege automation surface.

### Current workflow boundary

Static analysis can establish repository properties such as:

~~~text
required files
module names
Option policy
encoding and line endings
procedure structure
PtrSafe declarations
duplicate procedures
public API manifest
release-state markers
tracked-file hygiene
Markdown links
house-style formatter state
~~~

It cannot establish:

~~~text
VBE import success
VBA compile success
Excel object-model behavior
Windows API runtime behavior
Office bitness behavior
SDI window identity behavior
Ribbon active-window behavior
runtime cleanup
release workbook packaging
~~~

Do not present a green static workflow as proof that Excel executed or certified
the component.

### Workflow change controls

Changes to automation should preserve:

- least-privilege permissions;
- no secrets for ordinary static checks;
- review of every third-party action;
- minimal dependency installation;
- exact-SHA checkout for the event being evaluated;
- separation of testing from signing and publishing;
- sanitized logs and artifacts.

The current workflow references official GitHub actions by major-version tags.
Moving to immutable full-SHA action pinning can strengthen supply-chain control.
Do not claim immutable pinning until it is implemented.

---

## 🖥️ Future self-hosted Windows/Excel runner

Real Excel execution requires a Windows host with Office installed.

For a public repository, a persistent self-hosted runner is a materially
different security boundary from GitHub-hosted static CI.

Untrusted pull-request code must not execute automatically on a long-lived
workstation containing:

~~~text
Office credentials
personal files
browser sessions
signing certificates
release secrets
other repositories
network-mounted drives
developer credentials
~~~

### Required runner principles

A future Excel runner should use:

~~~text
least privilege
isolated Windows account
no personal data
no developer browser or session state
no unnecessary secrets
trusted-trigger policy
finite job timeout
deterministic Excel cleanup
orphan EXCEL.EXE cleanup
rebuildable or disposable hosts where practical
separate release-signing context
machine-readable completion output
exact commit-SHA binding
retained sanitized evidence
~~~

### Pull requests from forks

Do not automatically execute arbitrary fork pull-request code on a persistent
self-hosted Windows/Excel runner.

Safer approaches include:

~~~text
manual approval after source review
trusted branches only
workflow_dispatch for a reviewed SHA
ephemeral disposable Windows runners
quarantined runners with no credentials or sensitive network access
~~~

The trigger and trust policy must be documented before such a runner becomes a
required release control.

### Release secrets

A test runner should not automatically receive:

~~~text
release-signing private keys
high-privilege GitHub PATs
unrelated repository secrets
developer credentials
~~~

Testing and signing are separate trust stages. Compromise of a test job must not
automatically become compromise of the release-signing context.

---

## 🔐 Secret-handling rules

Never commit:

~~~text
personal access tokens
GitHub tokens
private signing keys
PFX, P12, or PVK files
passwords
API credentials
client secrets
private certificates
production connection strings
~~~

The repository .gitignore blocks common local secret-file formats and helps
prevent accidental additions.

.gitignore is not a security control.

A secret committed once must be considered compromised even if the commit is
later deleted.

If a credential is exposed:

1. revoke or rotate it immediately;
2. determine the scope of access;
3. remove it from the current tree;
4. inspect workflow, release, and account activity;
5. assess whether history cleanup is useful;
6. assume copies may remain in clones, caches, artifacts, logs, and forks.

Never repurpose production or client credentials for demo, regression, or
release-certification work.

---

## 🧾 Logging, diagnostics, and evidence

Diagnostics should contain enough information to reproduce a problem without
exposing unnecessary user or workstation data.

Prefer recording:

~~~text
procedure or stage
sanitized error number and description
FailureCount
sanitized FailureList entries
target scope
window counts and identity transitions
commit SHA
Excel build
Office bitness
Windows version
mandatory case counts
skipped and failure counts
cleanup result
sanitized runner metadata
~~~

Avoid publishing:

~~~text
worksheet data
client names
credentials
connection strings
private environment variables
signing-key paths
user profile contents
arbitrary workbook dumps
sensitive local paths
unsanitized workbook or window captions
~~~

Certification evidence must identify exact source without becoming a dump of
the workstation or workbook.

Generated evidence stored in the temporary folder remains executable-context
output. Review and sanitize it before attaching it to a release or public issue.

---

## 🧪 Regression harness security considerations

The regression harness can manipulate real Excel process and native window
state.

It can:

~~~text
create and close workbook windows
change the active window
show and hide managed UI
read and write title-bar frame styles
exercise Ribbon Excel 4 commands
change and restore ScreenUpdating
capture, restore, and clear snapshots
retain Excel Window references
inject deterministic failures
write certification text and JSON to the temporary directory
~~~

Run it:

- in a controlled workbook;
- in a dedicated Excel instance where practical;
- with unrelated workbooks closed;
- after saving all work;
- without a caller-owned component snapshot;
- with an accessible show-all recovery macro;
- with permission to restart Excel if cleanup becomes uncertain.

A passing failure count is not sufficient if:

~~~text
the runner did not finish
mandatory cases were skipped
cleanup failed
the wrong commit was tested
the wrong workbook or source was loaded
the output was manually transcribed incorrectly
the self-test destroyed pre-existing caller state
~~~

Release evidence must come from the exact source under review and include
completion, mandatory-case, failure, skip, and cleanup status.

---

## 📣 Disclosure coordination

Please avoid public disclosure while:

- exploitability is still being assessed;
- a release fix is being prepared;
- users have not had reasonable time to update;
- a credential or signing secret remains active;
- a malicious release asset remains downloadable;
- a vulnerable runner remains reachable.

The maintainer may ask for:

- additional environment detail;
- a sanitized reproduction;
- confirmation against a candidate fix;
- verification on another Office bitness or Excel build;
- a reasonable embargo period.

The project does not require a reporter to surrender ownership of their
research. The goal is to reduce preventable user harm.

---

## 🧭 Security review checklist for maintainers

For a security-sensitive VBA-EXCEL_UI change, review:

~~~text
[ ] Trust boundary stated
[ ] Caller-controlled input identified
[ ] UI hiding is not represented as access control
[ ] Application-level and window-level scopes are accurate
[ ] Ribbon command remains fixed and non-injectable
[ ] Ribbon active-window targeting assessed
[ ] Win32 and Win64 declarations reviewed
[ ] Long and LongPtr use reviewed
[ ] Valid-zero and GetLastError handling reviewed
[ ] Native handle lifetime and reuse assessed
[ ] Window object and hWnd pairing assessed
[ ] IsWindow is not treated as identity proof
[ ] Only owned frame-style bits can change
[ ] SetWindowPos flags preserve geometry and Z-order
[ ] Partial native-write and refresh-debt behavior assessed
[ ] Snapshot ownership and replacement assessed
[ ] Retained Window references are released
[ ] New, closed, and recreated window behavior assessed
[ ] Project-reset behavior assessed
[ ] ScreenUpdating baseline and cleanup assessed
[ ] FailureCount remains authoritative
[ ] Diagnostics are sanitized and non-secret
[ ] Original error evidence is preserved
[ ] Independent show-all recovery remains available
[ ] Regression or fault injection added where deterministic
[ ] Manual Excel and Windows validation recorded where required
[ ] Mandatory cases, skips, failures, and cleanup are explicit
[ ] Exact source SHA is bound to evidence
[ ] Source-to-artifact claims remain exact
[ ] Release-workbook impact assessed
[ ] Workflow permissions remain least privilege
[ ] External actions and dependencies reviewed
[ ] Self-hosted runner exposure assessed
[ ] Secrets and signing keys isolated from untrusted code
[ ] Public documentation updated
[ ] Private review or client material excluded
~~~

A security review should distinguish:

~~~text
code correctness
window and host-state integrity
operational safety
security impact
release trust
~~~

They overlap. They are not identical.

---

## 📚 Related policies and documentation

- [Project README](README.md)
- [Installation and Upgrade Guide](INSTALLATION.md)
- [Contributing Guidelines](CONTRIBUTING.md)
- [Code of Conduct](CODE_OF_CONDUCT.md)
- [Changelog](CHANGELOG.md)
- [Ribbon SDI Behavior](docs/RIBBON_SDI_BEHAVIOR.md)
- [MIT License](LICENSE)
- [Project Wiki](https://github.com/danielep71/VBA-EXCEL_UI/wiki)

---

## 👤 Maintainer

Maintained by **Daniele Penza**.

Private security reports:

~~~text
danielep71@gmail.com
~~~

---

<div align="center">

## 🛡️ Security principle

**Trust the source you run. Never treat hidden UI as access control. Keep Ribbon commands fixed, native style writes narrowly owned, snapshot state explicitly owned, and release evidence bound to the exact source. Keep untrusted code away from privileged runners and signing material.**

</div>
