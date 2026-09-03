<div align="center">

# 🔒 VBA-EXCEL_UI Security Policy

### Security, recovery, and trust boundaries for Excel and Windows UI state

[![Reporting](https://img.shields.io/badge/Reporting-Private-d97706?style=for-the-badge)](#reporting-a-vulnerability)
[![Support](https://img.shields.io/badge/Support-Latest_release-217346?style=for-the-badge)](#supported-versions)
[![Scope](https://img.shields.io/badge/Scope-Source_%7C_Releases_%7C_Automation-0969da?style=for-the-badge)](#security-scope)
[![Disclosure](https://img.shields.io/badge/Disclosure-Coordinated-6f42c1?style=for-the-badge)](#coordinated-disclosure)

<br>

**Protect users · Minimize exposure · Preserve evidence · Coordinate disclosure**

</div>

---

**VBA-EXCEL_UI** distributes plain-text VBA modules and macro-enabled demonstration artifacts that can change workbook-, window-, application-, Ribbon-, and Windows-native UI state.

This policy explains which versions receive security attention, how to report a
suspected vulnerability privately, what the project considers security-relevant,
and which trust boundaries remain the responsibility of users and host
organizations.

> [!IMPORTANT]
> A security policy does not make macros, workbooks, add-ins, source archives, or
> release artifacts inherently trustworthy. Establish provenance and apply your
> organization's security controls before enabling executable content.

---

<a id="security-model"></a>

## 🧭 Security model

The project assumes that:

- Microsoft Excel, the operating system, and the VBA runtime are trusted;
- the current user is authorized to open and run the workbook or add-in;
- macros are enabled only through an approved trust mechanism;
- project source or artifacts were obtained from an official channel; and
- the host workbook and other code already trusted in the Excel process are not
  malicious.

These are trust boundaries, not guarantees. VBA projects running in the same
Excel process are not isolated security sandboxes.

Hiding Excel interface elements is a presentation and workflow feature. It is not access control, authorization, segregation of duties, or a substitute for workbook, VBA-project, operating-system, or organizational security controls.

---

<a id="supported-versions"></a>

## 📦 Supported versions

| Source state | Security support |
|---|---|
| **Latest tagged functional release** | ✅ Supported |
| **Release candidate before publication** | ⚠️ Testing and best-effort remediation |
| **main** | ⚠️ Development code; best effort |
| **Older tagged releases** | ❌ Normally unsupported; upgrade first |
| **Modified copies, unofficial forks, or mirrors** | ❌ Unsupported unless the issue reproduces in official supported source |

If the project has not yet published a functional release, development code is
pre-release and no production version is security-supported.

Security fixes normally land on **main** and are included in a new tagged
release. Older releases are not normally patched in place unless the maintainer
states otherwise.

Reports must identify an exact release tag or full commit SHA. Descriptions such
as “latest” or “yesterday's main” are insufficient because branches change.

---

<a id="reporting-a-vulnerability"></a>

## 📣 Reporting a vulnerability

Do **not** disclose a suspected vulnerability in a public issue, discussion,
pull request, commit message, Wiki page, sample workbook, screenshot, or release
thread.

Use either private channel:

1. On the repository **Security** page, select **Report a vulnerability** when
   GitHub private vulnerability reporting is available.
2. Otherwise email **danielep71@gmail.com** with the subject:
   **Private security report — VBA-EXCEL_UI**.

Include the smallest amount of information needed to reproduce and assess the
issue:

| Evidence | Requested detail |
|---|---|
| 🧾 **Identity** | Repository, exact tag or full commit SHA, file, module, procedure, and artifact |
| 🖥️ **Environment** | Excel version/build, Office bitness, Windows version, locale, and deployment model |
| 🎯 **Impact** | Confidentiality, integrity, availability, code-execution, or supply-chain consequence |
| 🔬 **Reproduction** | Minimal steps and proof using synthetic data |
| 🧨 **Exploitability** | Preconditions, required trust, user interaction, affected scope, and persistence |
| 🛡️ **Mitigation** | Workaround or containment already tested, if any |
| 📎 **Evidence** | Sanitized logs, diagnostics, screenshots, hashes, or proof of concept |
| 🪟 **UI context** | Workbook and window identity, active window, affected surface, handle/style state, and multi-window conditions |
| 🔄 **Lifecycle** | Capture, apply, nested use, restore, emergency cleanup, process exit, and any residual state |
| 🧩 **Command path** | Public procedure, Ribbon callback, fixed Excel macro command, WinAPI call, or other path involved |

Do not send real client, employer, counterparty, student, production, or personal
workbooks. Remove credentials, tokens, personal data, internal paths,
connections, external links, document metadata, hidden names, cached values,
queries, and other unrelated content.

If a secret has been exposed, revoke or rotate it immediately before spending
time perfecting the report.

---

<a id="response-process"></a>

## ⏱️ Response process

This project is maintained by one person. Response times are best-effort rather than a contractual SLA.

The maintainer aims to:

| Stage | Target |
|---|---|
| **Acknowledgement** | Within 5 business days |
| **Initial scope and severity assessment** | Within 10 business days after sufficient evidence is available |
| **Progress update for an active investigation** | At least every 14 days |
| **Remediation and disclosure** | Proportionate to severity, exploitability, affected users, and validation needs |

The process normally includes reproducing the issue, determining affected
versions and artifacts, containing active risk, developing a fix, adding
regression or fault-injection evidence, validating in the relevant Excel and
Windows environment, and preparing a corrected release or advisory.

Targets may change when reproduction requires unavailable Office versions,
hardware, long-running behavior, third-party coordination, or sanitized evidence
from the reporter. Material delays will be communicated when practical.

Reporter credit can be included in an advisory or release notes when requested.
Anonymous credit is also acceptable.

---

<a id="security-triage"></a>

## 🎯 Security issue or ordinary defect?

When uncertain, report privately. The maintainer can reclassify a report safely.

Security reports include credible risks of:

- unintended code execution or crossing a documented trust boundary;
- unauthorized reading, modification, deletion, or disclosure of data;
- persistent or exploitable loss of availability;
- credential, token, signing-key, runner, or automation compromise;
- malicious, substituted, or misleading official release artifacts;
- validation or provenance bypasses that can represent an unsafe artifact as
  trusted; or
- a correctness defect deliberately exploitable to defeat a security,
  authorization, integrity, or control boundary.

An incorrect result, compatibility problem, bounded performance regression,
documentation error, or recoverable UI defect is normally an ordinary bug unless
it creates a concrete security impact.

### Severity guide

| Severity | Typical impact |
|---|---|
| **Critical** | Unintended code execution, exposed release credentials, compromised official artifacts, or broad unauthorized data access |
| **High** | Significant integrity/confidentiality loss, persistent host compromise, or practical supply-chain exploitation |
| **Moderate** | Bounded availability or integrity impact requiring meaningful preconditions |
| **Low** | Hardening weakness or limited impact without demonstrated exploitation |

Severity considers impact, exploitability, required privileges, user
interaction, affected versions, recoverability, and whether trusted malicious
VBA is already required.

---

<a id="security-scope"></a>

## 🛠️ Security scope

### In scope

- official source and committed executable or macro-enabled artifacts;
- official GitHub Release assets, archives, checksums, manifests, and provenance
  claims;
- repository-owned build, test, validation, packaging, and release tooling;
- GitHub Actions workflows, permissions, dependencies, and project-managed
  credentials;
- documented runtime integrations and trust boundaries; and
- security or integrity behavior introduced by this project's code.

### Project-specific risk surfaces

- **Trust-boundary confusion** — representing hidden UI as protection against users, macros, add-ins, automation, or file manipulation.
- **Process-wide effects** — state changes that escape the documented workbook/window scope or cannot be recovered safely.
- **Native window handling** — stale or recycled handles, incorrect pointer-width declarations, unsafe style masks, and failed rollback.
- **Command construction** — changing fixed Ribbon or Excel macro commands into user-controlled executable text.
- **Artifact integrity** — undisclosed macros, links, connections, payloads, or differences between source, tags, and demo workbooks.
- **Availability** — persistent UI loss, hangs, loops, crashes, or recovery failures beyond an ordinary bounded defect.

### Out of scope

- vulnerabilities in Microsoft Excel, Office, Windows, GitHub, Python, or the
  VBA runtime themselves;
- organization-controlled macro security, endpoint controls, access rights, or
  deployment policy;
- malicious VBA already trusted and running in the same Excel process;
- unrelated workbooks, add-ins, dependencies, or infrastructure;
- modified copies that do not reproduce the issue in official supported source;
- unofficial mirrors, repackaged binaries, or unsupported historical snapshots;
- lost or stolen user credentials not exposed by this project;
- social engineering unrelated to official project content; and
- ordinary defects without a concrete security impact.

Upstream vulnerabilities should be reported to the responsible vendor or
platform.

---

<a id="data-and-secrets"></a>

## 🔐 Data and secret handling

Never commit, upload, log, or attach:

- passwords, personal access tokens, API keys, signing keys, certificates, or
  connection strings;
- client, employer, counterparty, student, employee, or personal data;
- proprietary source, models, workbooks, market data, production extracts, or
  licensed vendor content;
- internal URLs, machine-specific paths, environment dumps, or unredacted
  screenshots; or
- proof-of-concept material beyond what is necessary to establish the issue.

Use synthetic data and a minimal reproduction. Excel files can carry sensitive
material outside visible cells, including document properties, defined names,
hidden sheets, VBA, cached values, Power Query data, external links, and
connections.

Repository secrets must be scoped to the smallest necessary workflow, protected
from untrusted pull-request code, excluded from logs and artifacts, and rotated
after suspected exposure.

---

<a id="supply-chain"></a>

## 📦 Supply-chain and release integrity

Trusted distribution is limited to the official repository and its GitHub
Releases page.

Maintainers should:

- review executable and macro-enabled artifacts before publication;
- pin third-party workflow actions to immutable commit SHAs;
- grant workflows the minimum required permissions;
- keep build, validation, signing, and publication responsibilities separated
  where practical;
- publish checksums, manifests, attestations, or signatures when the release
  process supports them; and
- document exactly what each piece of release evidence proves.

A checksum proves file identity. It does not prove that the file is safe, that
it was built from the stated source, or that it executed successfully in Excel.

Source hashes, artifact hashes, source-to-artifact provenance, Excel execution
evidence, and signing identity are distinct claims and must not be conflated.

---

<a id="automation"></a>

## 🤖 Repository automation and runners

Workflow code and configuration are security-sensitive.

- Do not expose secrets to pull requests from forks or other untrusted code.
- Do not run untrusted contributions on a persistent self-hosted Excel/Windows
  runner with repository, user, network, or signing credentials.
- Use ephemeral or isolated runners where practical.
- Clean workbooks, temporary files, Excel processes, credentials, and workspace
  state between jobs.
- Treat logs, screenshots, workbooks, test artifacts, and environment metadata
  as potentially sensitive.
- Review changes to workflow permissions, action pins, release jobs, dependency
  acquisition, and artifact upload/download paths as security changes.

Automation that writes repository content or publishes releases must have
explicit, least-privilege authorization.

---

<a id="safe-use"></a>

### Current automation surface

The repository currently contains three workflows:

~~~text
.github/workflows/static-checks.yml
.github/workflows/sync-label-colors.yml
.github/workflows/wiki-badges.yml
~~~

The static workflow:

- runs on a GitHub-hosted Ubuntu runner;
- checks exported source as text;
- uses the repository’s Python tooling;
- installs no project package dependencies;
- declares contents: read;
- does not require repository secrets;
- does not execute Excel or VBA;
- does not package or publish a release workbook.

The Wiki workflow checks the separately hosted Wiki against the root `VERSION`
file. It clones the public Wiki without credentials, runs on a weekly schedule
and on relevant repository changes, and cannot write to the Wiki.

The label-color workflow is the only current workflow with write permission. It
has `issues: write` solely to apply the repository's fixed label-color map, runs
manually or when its own workflow file changes, and neither reads nor publishes
release artifacts. Its broader permission must not be copied into static or
Wiki jobs.

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
workflow pin and tag-trigger policy
tracked-file hygiene
Markdown links
Markdown escape and private-review exclusion policy
Wiki-badge rule fixtures
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

Every third-party Action reference is pinned to a full 40-character commit SHA
with a trailing release-version comment. `tools/check_repo.py` enforces the pin
shape across every tracked workflow, and the static workflow also runs on `v*`
tag pushes so the tagged SHA receives its own repository check. A pin still
requires human verification against the upstream release when it is updated;
immutability does not prove that the selected upstream commit is trustworthy.

---


## ✅ Safe-use guidance

Users should:

- obtain source and demo artifacts only from the official repository or Releases page;
- treat macro-enabled workbooks as executable content and inspect them before enabling macros;
- preserve organization-approved macro security, Protected View, signing, and deployment controls;
- document process-wide effects and keep a trusted emergency UI-restoration path;
- test capture, nested use, restore, and cleanup in a controlled non-production workbook;
- sanitize workbooks and screenshots before sharing them; and
- review WinAPI, Excel macro-command, handle, workflow, and release-artifact changes as security-sensitive.

No numerical, statistical, pricing, timing, or UI result from this project is by
itself an authentication, authorization, access-control, cryptographic,
financial-advice, or safety-critical mechanism.

---

<a id="coordinated-disclosure"></a>

## 📣 Coordinated disclosure

Avoid public disclosure while exploitability is being assessed, a fix is being
prepared, affected users have not had reasonable time to update, an exposed
secret remains valid, or a malicious artifact or runner remains reachable.

The maintainer and reporter should agree a disclosure plan based on severity,
active exploitation, remediation complexity, availability of a workaround, and
the time needed to validate a corrected release.

The maintainer may ask for a sanitized reproduction, additional environment
detail, confirmation against a candidate fix, or a reasonable embargo. The
reporter does not surrender ownership of their research.

When remediation is available, the project may publish a GitHub Security
Advisory, corrected release, release-note entry, mitigation guidance, and credit
agreed with the reporter.

---

<a id="safe-harbor"></a>

## 🛡️ Good-faith research and safe harbor

Good-faith security research is welcome when it:

- stays within this project's source, artifacts, and documented integrations;
- avoids privacy violations, data destruction, service disruption, persistence,
  social engineering, and access to data beyond what is necessary;
- stops after establishing the minimum evidence required;
- reports the issue privately and promptly; and
- allows reasonable time for investigation and remediation.

The project will not initiate or recommend legal action solely for research
conducted in good faith and consistently with this policy. This statement does
not authorize testing of third-party systems or bind Microsoft, GitHub, an
employer, a client, or any other third party.

No paid bug bounty is offered unless the maintainer states otherwise in writing.

---

<a id="related-policies"></a>

## 📚 Related policies

- [Code of Conduct](CODE_OF_CONDUCT.md)
- [Contributing Guidelines](CONTRIBUTING.md)
- The repository README, license, release notes, and project documentation where
  present
- GitHub's platform security and acceptable-use policies

Conduct complaints and vulnerability reports are different. Use the Code of
Conduct for participant behavior and this policy for software risk.

---

<div align="center">

### Security principle

**Trust deliberately · Run minimally · Protect secrets · Preserve evidence · Disclose responsibly**

<br>

Maintained by **Daniele Penza**

</div>
