# 🚀 VBA Excel UI Release Guide

[![Release model: exact source](https://img.shields.io/badge/release-exact%20source-0969da)](#release-invariants)
[![Versioning: SemVer](https://img.shields.io/badge/versioning-SemVer-3f4551)](#versioning)
[![Evidence: required](https://img.shields.io/badge/evidence-required-success)](#evidence-record)
[![Security policy](https://img.shields.io/badge/security-policy-success)](SECURITY.md)
[![License: MIT](https://img.shields.io/badge/license-MIT-yellow.svg)](LICENSE)

This maintainer guide turns a reviewed commit into a traceable VBA Excel UI release. Source identity, validation, packaging, provenance, and publication must describe the same candidate.

> [!IMPORTANT]
> Certification must run in interactive Windows desktop Excel. Static checks alone cannot validate multi-window identity, title-bar integration, or restoration behavior.

## 🧭 Release profile

| Property | Requirement |
| --- | --- |
| Project maturity | Versioned Windows Excel UI library |
| Version source | [VERSION](VERSION) |
| Version scheme | Semantic Versioning |
| Tag format | Lower-case `vX.Y.Z` matching `VERSION` |
| Changelog | [CHANGELOG.md](CHANGELOG.md) |
| Installation contract | [INSTALLATION.md](INSTALLATION.md) |
| Vulnerability handling | [SECURITY.md](SECURITY.md) |
| License | [MIT](LICENSE) |

<a id="release-invariants"></a>
## 🔒 Release invariants

A release is valid only when all of these statements are true:

1. The candidate is identified by its full SHA and is reachable from protected `main`.
2. Release scope is frozen; unrelated work is deferred.
3. `VERSION`, changelog, source headers, package metadata, and documentation agree.
4. Static checks pass on the exact candidate.
5. VBA compiles and project-specific certification passes on that candidate.
6. Every binary artifact is built from that candidate, then tested as packaged.
7. Hashes and evidence bind artifacts to the candidate and tool environment.
8. The annotated lower-case tag points to the certified commit.
9. The GitHub Release uses that tag and the reviewed notes and assets.
10. Post-publication checks confirm that a new user can retrieve and validate it.

If an invariant is false, stop and repair the candidate. Never compensate by editing an already-tested artifact manually.

## 🗂️ Authority map

| Concern | Authoritative record |
| --- | --- |
| Current project version | `VERSION` |
| User-visible changes | `CHANGELOG.md` |
| Released source identity | Annotated Git tag and full commit SHA |
| Supported installation | `INSTALLATION.md` |
| Published binaries | GitHub Release assets |
| Integrity | SHA-256 hashes recorded with the release |
| Validation | Evidence record retained for the candidate |
| Vulnerability disclosure | `SECURITY.md` |

<a id="versioning"></a>
## 🧮 Versioning

Use `MAJOR.MINOR.PATCH` without a leading `v` inside `VERSION`.

- **MAJOR**: incompatible public API, behavior, file-format, or migration change.
- **MINOR**: backward-compatible capability.
- **PATCH**: backward-compatible correction or package/documentation fix included in the release.
- Use pre-release identifiers only when their support meaning is documented.

The Git tag adds the lower-case prefix: version `1.2.3` becomes annotated tag `v1.2.3`. Do not use upper-case `V`, moving tags, or a tag that differs from `VERSION`.

## ✅ Readiness review

- The runtime, title-bar, snapshot, and facade modules match the installation manifest.
- The candidate introduces no unreviewed Windows API or host-state side effects.
- Every planned item is merged or explicitly deferred.
- Compatibility and migration consequences are understood.
- Security-sensitive work has completed private handling where necessary.
- The working tree and exported VBA sources are reproducible.
- Maintainers and required reviewers are available.

Record which Windows and Office-bitness combinations were exercised; do not generalize beyond the tested matrix.

## 1. Freeze and identify the candidate

1. Update refs and start from current protected `main`.
2. Create the release-preparation branch required by repository policy.
3. Record the base and candidate full commit SHAs.
4. Stop feature work on that branch.
5. Review the complete diff from the previous release tag.
6. Confirm that generated and binary changes are intentional.

```bash
git fetch --tags --prune
git rev-parse HEAD
git status --short
git diff --stat <previous-tag>...HEAD
```

A dirty tree, unknown generated file, or unreviewed binary delta is blocking.

## 2. Synchronize version surfaces

Update every applicable surface in one reviewable change:

- `VERSION`
- `CHANGELOG.md`
- the four production VBA module headers
- test and demo version stamps
- README compatibility and installation examples

Search for the prior version and retain only intentional historical references. Never rewrite historical changelog sections or immutable evidence.

## 3. Finalize the changelog

Move relevant entries from **Unreleased** into a dated `[MAJOR.MINOR.PATCH] - YYYY-MM-DD` section.

- Describe user-visible behavior, not commit mechanics.
- Use Added, Changed, Deprecated, Removed, Fixed, and Security where applicable.
- Link issues or pull requests when they improve traceability.
- Do not claim an artifact, platform, or guarantee that was not certified.
- Keep an empty Unreleased section for future work.
- Add or verify the comparison link for the new version.

## 4. Verify documentation and installation

- Follow [INSTALLATION.md](INSTALLATION.md) from a clean environment.
- Verify source paths, import order, module names, prerequisites, and upgrades.
- Confirm README examples use the supported API and current version.
- Check [CONTRIBUTING.md](CONTRIBUTING.md), [SECURITY.md](SECURITY.md), and license links.
- Remove stale commands and promises beyond the tested support matrix.

## 5. Run static gates

- Run the repository Markdown, link, and exported-source checks.
- Run the repository formatter or formatting verification.
- Validate 32-bit and 64-bit declarations and scan for secrets or generated files.

Capture commands, tool versions, timestamps, and complete results. Rerun affected gates after any change.

## 6. Certify in Excel

- Run **Debug → Compile VBAProject** on the exact candidate.
- Run `Test_EXCEL_UI_RunReleaseCertification`.
- Exercise hide/show, snapshot/restore, multi-window identity, and title-bar behavior.
- Confirm cleanup restores host UI state after success and error paths.

Certification rules:

- Use a clean workbook or documented clean fixture.
- Import only candidate files.
- Record Excel version, Windows version, and Office bitness.
- Test the advertised environment matrix.
- Treat warnings, repairs, or unexplained numerical deltas as failures.
- If code changes after certification, restart static and Excel validation.

## 7. Build release artifacts

Planned outputs:

- Demo or certification `.xlsm` workbook, when published
- Source archive created by GitHub from the tag

For each artifact:

1. Start from a clean build location.
2. Use only candidate-controlled inputs.
3. Preserve required form and resource companions.
4. Compile before saving.
5. Exclude developer-only tests unless the artifact promises them.
6. Reopen and run the packaged smoke or regression test.
7. Record filename, size, and SHA-256.
8. Never edit the artifact after hashing.

```powershell
Get-FileHash -Algorithm SHA256 .\dist\<artifact>
```

```bash
sha256sum dist/<artifact>
```

Publish only artifacts promised by the installation guide.

## 8. Review and merge

Where policy requires a pull request, make these items easy to verify:

- intended version, previous tag, and candidate SHA;
- changelog and version synchronization;
- static and Excel results;
- artifact manifest and hashes;
- compatibility, migration, and security notes;
- remaining limitations.

Require configured checks and record the resulting `main` SHA. If the merge
changes the source commit identity, every pre-merge binary artifact and its hash
are stale—even when the resulting source tree appears unchanged. Before tagging:

1. check out the resulting `main` SHA;
2. rerun static and Excel certification on that exact commit;
3. rebuild every published binary artifact from that commit;
4. reopen and test each rebuilt artifact as packaged;
5. recalculate its size and SHA-256; and
6. replace the pre-merge artifact evidence with the resulting `main` evidence.

Do not tag or upload an artifact built from the pull-request head while claiming
the merge commit as its source identity.

## 9. Create the annotated tag

Tag only the certified commit on `main`.

```bash
git switch main
git pull --ff-only
git rev-parse HEAD
git tag -a vX.Y.Z -m "VBA Excel UI X.Y.Z"
git show --no-patch --decorate vX.Y.Z
git push origin vX.Y.Z
```

Before pushing, confirm the tag equals `VERSION`, targets the certified commit, and has a matching dated changelog section. Never delete and recreate a public tag to hide a mistake.

<a id="evidence-record"></a>
## 🧾 Evidence record

Retain at least:

- candidate commit SHA
- static and formatting results
- Excel version and bitness
- release-certification result
- artifact SHA-256 hashes
- version and tag
- previous release tag
- release-preparation pull request
- artifact filenames and sizes
- validation timestamps
- known deviations and approving reviewer

A release note summarizes evidence; it does not replace it.

## 10. Publish the GitHub Release

Create the release from the annotated tag and include:

1. User-facing summary and highlights.
2. Upgrade or migration notes.
3. Supported platform statement.
4. Known limitations.
5. Installation link.
6. Asset table with SHA-256 hashes.
7. Full changelog comparison link.
8. Security-reporting link.

Upload the already-hashed artifacts. Do not rebuild between tagging and upload.

## 11. Verify after publication

| Check | Result |
| --- | :---: |
| Tag resolves to the certified SHA | ☐ |
| `VERSION` and changelog match the tag | ☐ |
| Assets download and hashes match | ☐ |
| Installation links and examples work | ☐ |
| Packaged artifact opens and passes its smoke test | ☐ |
| Source archive contains expected release files | ☐ |
| Default branch is ready for the next Unreleased cycle | ☐ |

Do not announce broad availability until these checks pass.

## 🧯 Recovery

### Before tag publication

Fix the branch or pull request, update evidence, and rerun affected gates. Rebuilding invalidates the prior hash and packaged-test result.

### After tag, before GitHub Release

Pause publication and document the correction. Prefer a new patch candidate whenever the tag may have propagated.

### After public release

Do not silently replace assets or move the tag.

- Mark a dangerous release clearly and remove assets only for user safety.
- Publish a corrected patch with a new tag.
- Record the problem and remedy in the changelog.
- Use [SECURITY.md](SECURITY.md) for vulnerabilities.
- Preserve evidence explaining what users received.

## ☑️ Final maintainer checklist

| # | Gate | Done |
| ---: | --- | :---: |
| 1 | Scope frozen and diff reviewed | ☐ |
| 2 | Version surfaces synchronized | ☐ |
| 3 | Changelog finalized | ☐ |
| 4 | Installation verified cleanly | ☐ |
| 5 | Static checks pass | ☐ |
| 6 | Excel certification passes | ☐ |
| 7 | Artifacts built and packaged-tested | ☐ |
| 8 | Hashes and evidence recorded | ☐ |
| 9 | Required pull request merged | ☐ |
| 10 | Annotated tag targets certified main SHA | ☐ |
| 11 | GitHub Release published | ☐ |
| 12 | Post-publication checks pass | ☐ |

## 📚 Related documents

- [README.md](README.md) — project overview
- [INSTALLATION.md](INSTALLATION.md) — deployment and removal
- [CONTRIBUTING.md](CONTRIBUTING.md) — change and review workflow
- [CHANGELOG.md](CHANGELOG.md) — release history
- [SECURITY.md](SECURITY.md) — supported versions and private reporting
- [VERSION](VERSION) — authoritative version
- [LICENSE](LICENSE) — MIT license

---

**Release principle:** certify one exact source revision, build from it once, and publish only evidence-backed artifacts derived from it.
