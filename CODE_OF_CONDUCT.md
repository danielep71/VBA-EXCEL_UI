<div align="center">

# 🧭 VBA-EXCEL_UI Code of Conduct

### Respectful, evidence-led collaboration for Excel and Windows UI state

[![Applies to](https://img.shields.io/badge/Applies_to-Everyone-217346?style=for-the-badge)](#scope)
[![Spaces](https://img.shields.io/badge/Spaces-Code_%7C_Issues_%7C_PRs-0969da?style=for-the-badge)](#scope)
[![Standard](https://img.shields.io/badge/Standard-Respectful_%2B_Evidence--Led-6f42c1?style=for-the-badge)](#technical-collaboration)
[![Enforcement](https://img.shields.io/badge/Enforcement-Fair_%2B_Proportionate-d97706?style=for-the-badge)](#enforcement)

<br>

**Respect people · Challenge ideas with evidence · State uncertainty · Protect data**

</div>

---

**VBA-EXCEL_UI** is an open-source Excel/VBA project that captures, changes, and restores user-interface state across workbooks, windows, Excel sessions, and Windows configurations.

Technical rigor and respectful interaction are complementary requirements.
Neither excuses the absence of the other.

People should feel comfortable asking questions, reporting defects, challenging
assumptions, proposing safer or more accurate alternatives, and correcting an
earlier conclusion when new evidence emerges.

---

<a id="our-pledge"></a>

## 🤝 Our pledge

Everyone who participates through code, issues, pull requests, reviews,
documentation, examples, releases, the Wiki, or project discussion is expected
to help create a harassment-free experience for all.

That expectation applies regardless of age, body size, disability, ethnicity,
gender identity or expression, level or type of experience, nationality,
personal appearance, race, religion, socioeconomic status, sexual identity or
orientation, or any other personal characteristic unrelated to the
contribution.

We commit to acting and interacting in ways that support an open, welcoming,
diverse, inclusive, and healthy community.

---

<a id="expected-behavior"></a>

## ✅ Expected behavior

| Principle | Expected practice |
|---|---|
| 🤝 **Respect** | Assume good faith and address the work rather than the person. |
| 🎯 **Precision** | Distinguish observation, requirement, assumption, inference, hypothesis, and opinion. |
| 🧪 **Evidence** | Provide reproducible examples and relevant environment details where practical. |
| 🧭 **Transparency** | State uncertainty, limitations, conflicts of interest, and material dependencies. |
| 🔄 **Correction** | Acknowledge mistakes openly when better evidence changes the conclusion. |
| 🌱 **Inclusion** | Welcome contributors with different backgrounds and levels of expertise. |
| 🔐 **Stewardship** | Respect privacy, confidentiality, security, licensing, and intellectual-property boundaries. |
| 🧱 **Coherence** | Accept that maintainers may adopt, revise, defer, split, or decline a contribution. |

### Useful disagreement

A useful technical disagreement is specific and testable:

> "The observed result differs from the documented contract outside the stated
> tolerance. The exact version, environment, minimal reproduction, expected
> result, observed result, and independent reference are provided below."

A personal judgment is not testable:

> "This is wrong because the author does not understand the subject."

Only the first statement helps improve the project.

---

<a id="unacceptable-behavior"></a>

## 🚫 Unacceptable behavior

Unacceptable behavior includes:

- harassment, intimidation, discrimination, threats, or personal attacks;
- trolling, insulting or derogatory comments, and deliberately inflammatory
  language;
- unwelcome sexual attention or sexualized language or imagery;
- publishing private or confidential information without permission;
- deliberately misrepresenting results, sources, authorship, test evidence, or
  another participant's statements;
- fabricating, altering, or selectively presenting evidence to conceal a
  material limitation or contrary result;
- pressuring others to disclose employer, client, counterparty, student, or
  proprietary information;
- using credentials, reputation, job title, or academic status to silence a
  technical challenge rather than addressing its substance;
- spam, unrelated promotion, commercial solicitation, or sustained disruption;
- public disclosure of a suspected vulnerability before reasonable coordinated
  remediation;
- repeated disruption after a maintainer has asked participants to stop; and
- retaliation against a reporter, witness, or participant in an investigation.

Disagreement is allowed. Abuse is not.

---

<a id="technical-collaboration"></a>

## 🧪 Technical collaboration

UI defects are often state-dependent. Reports should identify the affected workbook and window identity, the UI surface, and the lifecycle stage at which behavior diverged.

Where relevant, technical reports and review comments should identify:

| Evidence | Expected information |
|---|---|
| 🧾 **Identity** | Repository release, tag, branch or commit; affected file, procedure, API or artifact |
| 🖥️ **Environment** | Excel and Office version/build, 32-bit or 64-bit, Windows version, and relevant locale or host settings |
| 🔬 **Reproduction** | Smallest deterministic input and exact steps needed to reproduce the behavior |
| 🎯 **Comparison** | Expected and observed behavior, error or diagnostic output, and the acceptance criterion |
| ✅ **Validation** | Relevant automated, static, manual or Excel regression evidence and whether the run completed |
| ⚠️ **Boundary** | Assumptions, limitations, untested configurations, uncertainty, and what cannot yet be concluded |
| 🪟 **UI context** | Workbook and window identity, active window, affected UI surface, multi-window state, and relevant Excel global state |
| 🔄 **Lifecycle** | Capture, apply, nested use, restore, emergency cleanup, or process-exit stage |
| 🛡️ **Safety result** | State changed, state preserved, state refused, partial success, cleanup result, and any residual effect |

Prefer explicit evidence classifications:

| Classification | Meaning |
|---|---|
| **Observed** | Reproduced directly |
| **Derived** | Follows from stated inputs and rules |
| **Inferred** | Best explanation, but not independently proven |
| **Expected** | Required by the documented contract |
| **Environment-specific** | Verified only in the stated host configuration |
| **Unverified** | Plausible, but evidence is incomplete |

Do not present a plausible inference as a verified fact. A screenshot may
illustrate a result, but it does not replace the inputs, environment, contract,
reference, and steps needed to reproduce the claim.

Project-specific expectations:

- Distinguish workbook-level, window-level, application-wide, and Windows-native state.
- Do not describe a restore as successful without checking the relevant postcondition and cleanup result.
- Use sanitized workbooks and screenshots; UI evidence can expose workbook names, file paths, users, connections, and other private data.

---

<a id="data-and-confidentiality"></a>

## 🔐 Data, privacy, and confidentiality

Do not upload confidential or restricted material to demonstrate a defect or
support a contribution. This includes:

- client, employer, counterparty, student, or personal data;
- credentials, tokens, signing material, connection strings, or internal URLs;
- proprietary code, models, workbooks, market data, business assumptions, or
  production extracts;
- licensed vendor content that cannot be redistributed; and
- files or examples that the contributor is not authorized to share.

Use the smallest synthetic example that preserves the relevant behavior.

Excel workbooks can contain hidden names, external links, connections, cached
values, queries, metadata, comments, hidden sheets, and VBA that are not visible
on the active sheet. Sanitize and inspect every reproduction before uploading
it.

---

<a id="security"></a>

## 🔑 Security reports

Suspected vulnerabilities must follow [SECURITY.md](SECURITY.md). Do not
publish exploit details, credentials, private keys, or a working proof of
concept in a public issue before coordinated remediation.

The Code of Conduct reporting channel is for participant behavior. A security
report concerns software risk. If an incident involves both, use the private
channel and make that clear.

---

<a id="scope"></a>

## 🌐 Scope

This Code of Conduct applies to:

- source code, committed artifacts, documentation, examples, and releases;
- issues, pull requests, reviews, comments, discussions, and the Wiki;
- project-related email and private communication between participants; and
- public spaces where someone represents the project or its community.

It applies to maintainers, contributors, reviewers, users, and visitors alike.

Project representation includes using an official account, speaking on behalf
of the project, presenting oneself as a maintainer or contributor in a
project-related forum, or moderating a project discussion.

---

<a id="reporting"></a>

## 📣 Reporting unacceptable behavior

Report unacceptable behavior **privately** to the maintainer:

**danielep71@gmail.com**

Do not publish sensitive personal information in a public issue.

Where available, include what happened, where and approximately when it
happened, relevant links or screenshots, whether the behavior is ongoing,
whether another participant witnessed it, and any immediate safety, privacy, or
confidentiality concern.

Reports will be reviewed as promptly, fairly, and discreetly as reasonably
possible. Information will be shared only as needed to understand the report,
protect participants, enforce this policy, or comply with applicable platform
or legal requirements.

A good-faith report is not misconduct merely because the maintainer ultimately
concludes that no violation occurred.

---

<a id="enforcement"></a>

## ⚖️ Enforcement

The maintainer is responsible for clarifying and enforcing this Code of Conduct
and may remove, edit, or reject comments, commits, code, issues, or other
contributions that are inconsistent with it.

Responses depend on seriousness, frequency, context, prior behavior, and risk.
They may include:

1. clarification or a private reminder;
2. a formal warning and conditions for continued participation;
3. editing or removing project content;
4. closing or locking a discussion;
5. rejecting or reverting a contribution;
6. temporary restriction from project participation;
7. permanent exclusion from project spaces; or
8. escalation to GitHub or another relevant platform.

Enforcement aims to be fair, proportionate, consistent, protective of
participants, and protective of the technical record. Retaliation against
anyone who reports a concern or participates in its review is itself a
violation.

---

## 🧩 Conflicts of interest

Disclose a material interest when it could reasonably affect technical review.

Examples include ownership of a competing implementation, commercial interest
in a dependency or benchmark, employer or client restrictions, uncertainty
about the origin or license of submitted material, or reviewing one's own work
under another identity.

A conflict is not automatically disqualifying. Undisclosed material influence
is the concern.

---

## 📜 Source and licensing integrity

Contributors must have the right to submit every code fragment, document,
screenshot, workbook, dataset, benchmark, image, and numerical reference they
provide.

Identify the source and license of adapted material. Do not submit proprietary
code, incompatible licensed content, confidential screenshots, selectively
edited evidence, or generated material whose provenance and right of use cannot
be established.

A file digest proves file identity. It does not by itself prove authorship,
source provenance, reproducibility, or execution correctness.

---

## 🧱 Maintainer decisions

A maintainer may decline a contribution even when it is technically valid.
Reasons may include scope, compatibility, maintenance burden, testability,
platform risk, API stability, duplication, architecture, or release timing.

A declined contribution is not a judgment about the contributor. When
practical, the technical reason should be recorded. Participants may challenge
a decision respectfully with new evidence; repeatedly reopening the same
argument without new evidence is not constructive.

---

## 🙏 Project scale and response expectations

This project is maintained by one person. That affects response capacity,
not the seriousness of this policy.

Response times are best-effort. Complex reports may take longer when they
require a particular Office configuration, Windows behavior, clean Excel
process, long-running test, or manual workbook validation.

Reasonable delay is not dismissal. Repeatedly demanding immediate action is not
a substitute for technical evidence.

GitHub's platform policies and community standards also apply where relevant.

---

## 📜 Attribution

This Code of Conduct is informed by and adapted from the
[Contributor Covenant, version 2.1](https://www.contributor-covenant.org/version/2/1/code_of_conduct.html).
Its project-specific technical and data-handling provisions are maintained for
this repository.

---

<div align="center">

### Practical principle

**Be precise about the work · Be generous toward the person · Show the evidence · State the boundary · Protect the data**

<br>

Maintained by **Daniele Penza**

</div>
