<div align="center">

# 🧭 Code of Conduct

### Respectful, evidence-led technical collaboration for VBA-EXCEL_UI

[![Applies to](https://img.shields.io/badge/Applies_to-Everyone-217346?style=for-the-badge)](#scope)
[![Spaces](https://img.shields.io/badge/Spaces-Issues_PRs_Wiki_Releases-0969da?style=for-the-badge)](#scope)
[![Standard](https://img.shields.io/badge/Standard-Respectful_%2B_Evidence--Led-6f42c1?style=for-the-badge)](#technical-discussion-standards)
[![Enforcement](https://img.shields.io/badge/Enforcement-Maintainer-d97706?style=for-the-badge)](#enforcement)

<br>

**Technical rigor · Respectful disagreement · Reproducible evidence ·
UI-state safety · Privacy-aware collaboration**

</div>

---

**VBA-EXCEL_UI** is a focused open-source Excel/VBA component for controlling
and restoring Excel user-interface state.

This Code of Conduct exists to keep interaction around the project respectful,
technical, constructive, and welcoming—especially when behavior depends on
Excel's Single Document Interface, the active workbook window, Ribbon
availability, Windows handles, Office bitness, host policy, or UI state that is
difficult to reproduce safely.

People should feel comfortable:

- reporting defects;
- asking basic or advanced VBA questions;
- challenging an implementation or release claim;
- proposing safer or simpler alternatives;
- identifying behavior that differs by Excel build, bitness, policy, or window
  state;
- saying that an earlier assumption, test, document, or implementation decision
  was wrong;
- contributing even when they are unfamiliar with the project's conventions.

A technically demanding project benefits from disagreement.

It does not benefit from hostility.

---

<a id="our-pledge"></a>

## 🤝 Our pledge

Everyone who participates—by opening an issue, submitting a pull request,
commenting, reviewing, editing documentation or the Wiki, discussing a release,
or representing the project elsewhere—is expected to help create a
harassment-free experience for all.

That expectation applies regardless of:

- experience level;
- professional or academic background;
- age;
- disability;
- ethnicity;
- gender identity or expression;
- nationality;
- race;
- religion;
- sexual orientation;
- socioeconomic status;
- or any other personal characteristic unrelated to the technical contribution.

Technical rigor and respectful interaction are complementary requirements.

Neither excuses the absence of the other.

---

<a id="expected-behavior"></a>

## ✅ Expected behavior

Participants are expected to:

- be respectful and assume good faith unless evidence shows otherwise;
- focus criticism on code, behavior, documentation, tests, architecture,
  evidence, or process rather than on the person who produced them;
- distinguish **observed fact**, **test result**, **inference**, **hypothesis**,
  and **platform limitation**;
- describe Excel/VBA and Windows UI behavior precisely;
- provide reproduction steps, logs, screenshots, certification output, or
  minimal examples where practical;
- state the exact source and relevant host conditions when making compatibility
  or restoration claims;
- acknowledge uncertainty rather than presenting an assumption as verified;
- correct mistakes openly when new evidence changes the conclusion;
- give and receive constructive review comments professionally;
- respect privacy, confidentiality, client restrictions, and security
  boundaries;
- help newcomers understand the repository's workflow and vocabulary;
- allow maintainers reasonable time to investigate host-specific or
  multi-window behavior;
- respect that a contribution may be adopted, adapted, deferred, divided into
  smaller changes, or declined to preserve correctness, compatibility,
  maintainability, or release discipline.

### Useful disagreement

A useful technical disagreement is specific enough that another person can test
it:

> "The snapshot was captured with workbook window A active. After activating
> window B, restoration changed B's title-bar state while A remained unchanged.
> The exact commit, Excel build, Office bitness, window sequence, and regression
> output are below."

That can be investigated.

### Unhelpful disagreement

A personal judgment cannot be tested:

> "This code is useless because the author does not understand Windows."

Both statements may arise from frustration with the same problem.

Only the first helps fix it.

---

<a id="unacceptable-behavior"></a>

## 🚫 Unacceptable behavior

Unacceptable behavior includes:

- personal attacks, insults, ridicule, or derogatory comments;
- harassment in public or private;
- discriminatory, demeaning, or sexualized language or imagery;
- threats, intimidation, or encouragement of violence;
- publishing another person's private information without permission;
- deliberate misrepresentation of another participant's work or statements;
- knowingly fabricating test output, screenshots, provenance evidence,
  reproduction steps, or implementation claims;
- presenting a partial, interrupted, stale, or source-mismatched test run as
  complete release evidence;
- hiding a known limitation while making a compatibility, restoration, or
  release claim;
- sustained disruption of technical discussion;
- repeated bad-faith argument after the technical decision and evidence have
  been explained;
- spam, unrelated promotion, or commercial solicitation;
- attempts to pressure maintainers into unsafe disclosure, unverifiable claims,
  or an unverified release;
- public disclosure of a suspected vulnerability before reasonable coordinated
  remediation;
- retaliation against someone who reports misconduct, a security concern, an
  evidence-integrity concern, or a technical failure.

Disagreement is allowed.

Abuse is not.

---

<a id="technical-discussion-standards"></a>

## 🧪 Technical discussion standards

VBA-EXCEL_UI interacts with surfaces whose behavior can depend on the host
process, active Excel window, target scope, and Windows frame identity.

Relevant surfaces include:

```text
Application.CommandBars("Ribbon")
Application.DisplayStatusBar
Application.StatusBar
Application.Hwnd
Excel.Window.hWnd
DisplayHeadings
DisplayGridlines
DisplayWorkbookTabs
DisplayHorizontalScrollBar
DisplayVerticalScrollBar
WinAPI frame-style reads and writes
snapshot capture and restoration
quiet-update ownership and cleanup
```

A technical report is therefore more useful when its environment and UI state
are explicit.

### For runtime and restoration behavior

Where relevant, include:

- exact repository tag, branch, or commit SHA;
- Excel version and build;
- Office 32-bit or 64-bit;
- Windows version;
- number of open workbook windows;
- which Excel window was active at capture, mutation, and restoration;
- the requested `UIWindowTargetScope`;
- whether a snapshot already existed;
- the initial state of each affected UI surface;
- whether the Ribbon state was readable and writable;
- relevant window/hWnd identity observations;
- reproduction steps;
- observed behavior;
- expected behavior;
- regression evidence, logs, screenshots, or a minimal workbook when safe to
  share.

Do not infer window identity merely because a handle remains non-zero, a style
value matches, or the currently active window looks equivalent to the captured
one.

Do not infer successful restoration merely because a write call did not raise.
Where a readable postcondition exists, report the achieved state.

Where the platform makes a fact unavailable or impractical to verify, say so and
classify the evidence appropriately:

```text
automated static check
manual Excel regression
release certification
source inspection
Windows / VBA platform contract
host characterization
unresolved hypothesis
```

That distinction is part of this project's quality standard.

---

## ✅ Regression and certification evidence

For production changes, contributors should normally compile first:

```text
VBA Editor → Debug → Compile VBAProject
```

The interactive regression runner is:

```vb
Test_EXCEL_UI_RunAll
```

The release-certification runner is:

```vb
Test_EXCEL_UI_RunReleaseCertification
```

`Test_EXCEL_UI_RunAll` is useful during development but is not a substitute for
release certification.

A useful report identifies at least:

```text
runner
completion status
cases / units observed
failures
Excel version and build
Office bitness
Windows version
commit SHA
```

Machine-readable JSON or text evidence is preferable to manually transcribed
counts when it is available and safe to share.

Evidence must still be interpreted honestly:

- a passing count does not prove that every expected mandatory case ran unless
  the expected and observed inventories are both established;
- a file naming an Excel build does not by itself prove which source was loaded;
- an earlier run does not certify a later commit;
- a characterization probe is not a pass/fail regression;
- static checks do not substitute for importing, compiling, and executing the
  VBA in Excel;
- source inspection does not substitute for runtime evidence when the claim is
  about actual host behavior.

Before posting evidence publicly, inspect it for personal paths, workbook names,
usernames, client references, or other information that should remain private.

---

## 🪟 Excel UI-state and ownership discussions

This component can change process-wide and window-specific Excel state.

Those changes may affect more than the calling macro.

When discussing UI behavior, distinguish at least:

```text
captured baseline
baseline validity
captured Excel Window object
captured native hWnd
current active window
requested target scope
per-window state
process-wide state
successful write
verified achieved state
snapshot ownership
cleanup and emergency recovery
```

Do not describe a UI mutation as harmless merely because the calling procedure
is local to one workbook.

When reporting a restoration problem, include:

- the initial UI state;
- the capture and restore sequence;
- which windows were opened, activated, or closed;
- whether the native handle changed or became invalid;
- whether another snapshot or state owner already existed;
- whether cleanup completed normally or through recovery;
- which state was restored, left unchanged, or reported as unknown.

The project prefers:

```text
explicit state ownership
stable object/native identity pairing
verified postconditions
reported inability to restore
```

over:

```text
implicit global side effects
active-window assumptions
success inferred from no exception
guessing a replacement state
```

Technical discussion should preserve those distinctions.

---

## 🔍 Review standards

Review comments should be actionable whenever possible.

A strong review comment identifies:

1. **where** the concern exists;
2. **what** behavior or invariant is at risk;
3. **why** it matters;
4. **what evidence** supports the concern;
5. whether the requested change is:
   - required;
   - recommended;
   - optional;
6. whether the concern affects:
   - correctness;
   - SDI identity safety;
   - UI restoration;
   - compatibility;
   - diagnostics;
   - test validity;
   - documentation;
   - release provenance;
   - security;
   - maintainability;
   - style.

Example:

> "`UI_ResetExcelUIToSnapshot` retains a Window object and a native handle, but
> the restore path proves only that each remains independently usable. It does
> not prove that the retained Window currently owns the captured hWnd. A closed
> and recreated window can therefore redirect restoration if handle reuse is
> not rejected."

This is preferable to:

> "The snapshot code looks unsafe."

Reviewers should distinguish a blocking correctness defect from an optional
refactoring preference.

Contributors should not mark a concern resolved merely because code changed.
The relevant implementation, regression, documentation, and release evidence
must support closure.

---

## 📦 Source-first and release-evidence discussions

The repository is source-first.

Authoritative implementation lives in the four exported production modules:

```text
src/M_EXCEL_UI.bas
src/M_EXCEL_UI_RUNTIME.bas
src/M_EXCEL_UI_SNAPSHOT.bas
src/M_EXCEL_UI_TITLEBAR.bas
```

The demo workbook distributed through GitHub Releases is a convenience artifact,
not the authoritative source tree.

When discussing a release, distinguish:

```text
tag / commit identity
Git tree identity
source hashes
static-check evidence
Excel regression evidence
mandatory case inventory
Office bitness and build
demo-workbook digest
source-to-workbook provenance
```

A SHA-256 digest proves the identity of a file.

It does not, by itself, prove how that file was produced or which VBA source it
contains.

Do not claim reproducibility, execution certification, source provenance, or
issue closure beyond what the public evidence actually establishes.

Private assessments, audit material, or review artifacts must not be published
without the owner's explicit authorization. Public issue descriptions and
release records should be self-contained and must not require access to private
material.

---

<a id="privacy-and-confidentiality"></a>

## 🔒 Privacy, confidentiality, and safe reproductions

Do not upload confidential business material merely to demonstrate a UI,
window-identity, restoration, or Ribbon problem.

That includes:

- client workbooks;
- proprietary VBA;
- credentials;
- private signing keys;
- personal data;
- internal URLs;
- connection strings;
- production data extracts;
- private reviews, audit reports, or internal assessments;
- non-public add-ins or modules you are not authorized to distribute.

Create a sanitized minimal reproduction instead.

Excel files can contain more than visible worksheet values, including:

```text
defined names
external links
connection strings
Power Query metadata
cached values
comments
document properties
hidden and very-hidden sheets
VBA
```

A workbook that appears anonymized may still disclose information elsewhere in
the package.

If a private reproduction is genuinely necessary, coordinate privately with the
maintainer before sharing it. Do not assume that sending a private file grants
permission to publish, redistribute, quote, or commit it.

---

## 🔑 Security issues are different from ordinary bugs

Suspected vulnerabilities should follow [SECURITY.md](SECURITY.md).

Do **not** publish exploit details, credentials, private keys, or a working proof
of concept in a normal public issue before coordinated disclosure when doing so
could expose users or systems.

The Code of Conduct reporting channel is for participant behavior.

A security report is about software risk.

If an incident involves both, use the private channel and say so.

---

<a id="scope"></a>

## 🛠️ Scope

This Code of Conduct applies to:

- this GitHub repository;
- issues;
- pull requests;
- review comments;
- GitHub Discussions, if enabled;
- release threads;
- the Wiki;
- project-related email;
- project-related private communication between participants;
- public spaces where someone is representing the project.

Examples of project representation include:

- speaking on behalf of the project;
- using an official project account;
- presenting oneself as a maintainer or contributor in a project-related forum;
- moderating a project discussion.

The standards apply to both maintainers and contributors.

---

<a id="reporting-unacceptable-behavior"></a>

## 📣 Reporting unacceptable behavior

Report unacceptable behavior **privately** to the maintainer:

```text
danielep71@gmail.com
```

Where available, include:

- what happened;
- where it happened;
- dates or approximate times;
- links, screenshots, or quoted text;
- whether the behavior is ongoing;
- whether another participant witnessed it;
- any immediate safety, privacy, or confidentiality concern.

Do not post sensitive personal information in a public issue.

Reports will be handled as discreetly as reasonably possible.

Information will be shared only as needed to:

```text
understand the report
protect participants
enforce this policy
comply with platform or legal requirements where applicable
```

A good-faith report will not be treated as misconduct merely because the
maintainer ultimately concludes that no violation occurred.

Security vulnerabilities should instead follow
[SECURITY.md](SECURITY.md).

---

<a id="enforcement"></a>

## ⚖️ Enforcement

The maintainer is responsible for interpreting and enforcing this Code of
Conduct.

Responses depend on seriousness, frequency, context, prior behavior, and risk to
participants or the project.

Possible actions include:

1. clarification or a private reminder;
2. a formal warning;
3. editing or removing comments or contributions;
4. closing or locking a discussion;
5. rejecting or reverting a contribution;
6. temporary restriction from project participation;
7. permanent blocking;
8. escalation to GitHub or another relevant platform.

Enforcement aims to be:

```text
proportionate
consistent
documented where appropriate
protective of participants
protective of the technical record
```

Retaliation against a reporter, witness, or participant in an investigation is
itself a violation.

---

## 🧩 Conflicts of interest

Participants should disclose a material conflict when it could reasonably affect
technical review.

Examples:

- ownership of a competing implementation;
- commercial interest in a dependency, integration, or recommendation being
  proposed;
- employment or client restrictions that materially limit what can be
  disclosed;
- inability to establish the origin or license of copied code;
- evaluating a contribution one personally authored under another identity or
  organization.

A conflict is not automatically disqualifying.

Undisclosed material influence is the problem.

---

## 📜 Source and licensing integrity

Contributors must have the right to submit the code, documentation, screenshots,
test data, or other material they provide.

Do not submit:

- proprietary code copied from an employer or client;
- code with an incompatible license without clear attribution and discussion;
- screenshots containing confidential data;
- evidence altered or selectively presented to hide contrary results;
- generated code presented as independently authored when its provenance or
  license is uncertain;
- binary workbook changes that bypass the repository's source-first
  architecture.

If material was adapted from another source, identify that source and its license
clearly enough for review.

---

## 🧱 Maintainer decisions

A maintainer may decline a contribution even when it is technically valid.

Reasons can include:

```text
scope
compatibility
maintenance burden
testability
platform risk
API stability
UI-state safety
release timing
duplication
architectural direction
```

A declined contribution is not a judgment about the contributor.

When practical, the technical reason should be recorded so that the same design
question does not need to be rediscovered repeatedly.

Participants may challenge a decision respectfully with new evidence.

Repeatedly reopening the same argument without new evidence is not constructive.

---

## 🙏 Project scale and response expectations

VBA-EXCEL_UI is maintained by one person.

That affects response capacity, not the seriousness of this policy.

Response times are best-effort.

Complex reports may take longer when they require:

- a particular Office bitness or Excel build;
- two or more workbook windows;
- a clean Excel process;
- a specific Ribbon configuration or policy;
- fault-injection behavior;
- window closure and handle-reuse scenarios;
- multi-monitor behavior;
- a second Excel host or virtual machine;
- Windows API behavior that cannot be reproduced in hosted Linux CI.

Reasonable delay is not dismissal.

Repeatedly demanding immediate action is not an acceptable substitute for
technical evidence.

Where appropriate, GitHub's platform policies and community standards also
apply.

---

## 🧭 Practical principle

The standard for this project can be summarized as:

```text
be precise about the software
be honest about the evidence
be generous toward the person
state the uncertainty
protect UI state
protect user data
```

That is the environment in which difficult Excel/VBA UI-state and
window-identity problems are most likely to be solved well.

---

## 👤 Maintainer

Maintained by **Daniele Penza**.
