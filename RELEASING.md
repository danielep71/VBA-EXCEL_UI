# 🚀 VBA Excel UI Release Guide

[![Release model: exact source](https://img.shields.io/badge/release-exact%20source-0969da)](#release-invariants)
[![Versioning: SemVer](https://img.shields.io/badge/versioning-SemVer-3f4551)](#versioning)
[![Evidence: required](https://img.shields.io/badge/evidence-required-success)](#evidence-record)
[![Security policy](https://img.shields.io/badge/security-policy-success)](SECURITY.md)
[![License: MIT](https://img.shields.io/badge/license-MIT-yellow.svg)](LICENSE)

This guide describes how I turn a reviewed commit into a traceable VBA Excel UI release. Source identity, validation, packaging, provenance, and publication must describe the same candidate.

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
| Release notes | The version's section of `CHANGELOG.md` at the tagged commit |
| Integrity | `SHA256SUMS`, published with the release assets |
| Validation | Evidence record retained for the candidate |
| Vulnerability disclosure | `SECURITY.md` |

<a id="automation"></a>
## 🤖 Automated and manual steps

Pushing a `vX.Y.Z` tag starts the **Release** workflow
(`.github/workflows/release.yml`). It makes every decision with
`tools/release.py`, which you can also run locally:

| Step | `tools/release.py` | What it does |
| --- | --- | --- |
| Verify | `verify vX.Y.Z --repository danielep71/VBA-EXCEL_UI` | Checks that the tag is annotated and equals `v` + `VERSION` at the tagged commit, that `CHANGELOG.md` there has a dated section for the version, that the commit is reachable from `origin/main`, and that the required check **Repository and module checks** concluded success on that exact commit |
| Notes | `notes vX.Y.Z` | Prints that changelog section, which becomes the release body |
| Classify | `is-prerelease vX.Y.Z` | Prints `true` when the version has a pre-release suffix, such as `1.2.0-rc.1`, and `false` otherwise |
| Bundle | `bundle vX.Y.Z --out DIR` | Writes `git archive` output of the tag as `VBA-EXCEL_UI-X.Y.Z.zip` and `VBA-EXCEL_UI-X.Y.Z.tar.gz`, plus `SHA256SUMS`, and refuses an archive that still holds `AGENTS.md` or a path under `.github/`, two of the paths `.gitattributes` marks `export-ignore` |
| Publish | (workflow only) | Creates the GitHub Release from the tag with `gh release create`, using the notes and the bundle, and marks it as a pre-release when `is-prerelease` prints `true` |
| Verify after publication | `verify-published vX.Y.Z --repository danielep71/VBA-EXCEL_UI --dir DIR [--asset NAME=SHA256]` | Checks that the release exists and is not a draft, that its pre-release flag matches the version, that its body equals the notes, and that its assets are exactly the files in `SHA256SUMS`, that file, and each asset declared with `--asset`, all with matching SHA-256 values |

The verify and bundle jobs only read the repository. The publish job runs only
after both have passed, only on a tag push, and is the one job with
`contents: write`. Without `--repository`, `verify` checks everything except the
required check and reports `INCOMPLETE`.

To check an existing tag without publishing anything, run the workflow by hand
(**Actions → Release → Run workflow**) with the tag as input: it runs verify
and notes only.

These steps stay manual, and each must be complete before the tag is pushed:

- scope, version surfaces, the changelog and the installation guide (sections 1–4);
- Excel certification on the exact candidate (section 6) and the
  [evidence record](#evidence-record), which no workflow can produce;
- any workbook artifact (section 7), which the workflow neither builds nor
  uploads, and the `verify-published` run that includes it (section 11);
- the merge into `main` (section 8), and creating and verifying the annotated
  tag (section 9);
- the post-publication checks the workflow cannot run (section 11).

<a id="versioning"></a>
## 🧮 Versioning

Use `MAJOR.MINOR.PATCH` without a leading `v` inside `VERSION`.

- **MAJOR**: incompatible public API, behavior, file-format, or migration change.
- **MINOR**: backward-compatible capability.
- **PATCH**: backward-compatible correction or package/documentation fix included in the release.
- Use pre-release identifiers only when their support meaning is documented.

The Git tag adds the lower-case prefix: version `1.2.3` becomes annotated tag `v1.2.3`. Do not use upper-case `V`, moving tags, or a tag that differs from `VERSION`.

`python3 tools/check_repo.py` enforces the changelog structure this depends on.
`## [Unreleased]` must be the first version section. Every release heading must
read `## [X.Y.Z] - YYYY-MM-DD` with a real calendar date, newest release first.
Every release needs a link reference to this repository, either
`compare/v<previous>...v<this>` or `releases/tag/v<this>`. `VERSION` must equal
the newest release heading, so I bump `VERSION` and date the changelog section
in the same change.

## ✅ Readiness review

- The runtime, title-bar, snapshot, and facade modules match the installation manifest.
- The candidate introduces no unreviewed Windows API or host-state side effects.
- Every planned item is merged or explicitly deferred.
- Compatibility and migration consequences are understood.
- Security-sensitive work has completed private handling where necessary.
- The working tree and exported VBA sources are reproducible.
- I am available for the release, together with any required reviewers.

I certify releases on 64-bit Office only; 32-bit Office is not certified ([#51](https://github.com/danielep71/VBA-EXCEL_UI/issues/51)). Record which Windows versions and Excel builds were exercised; do not generalize beyond them.

## 1. Freeze and identify the candidate

1. Update refs. If the release branch `release/vX.Y.Z` is not open yet, create
   it from current protected `main`; otherwise use the open branch. When you
   open it, add its `target-branch: release/vX.Y.Z` entries to
   `.github/dependabot.yml` on `main`, one per ecosystem. `VERSION` and those
   entries move together: `X.Y.Z` must be the next patch, minor or major
   release after `VERSION`, and the change that bumps `VERSION` to `X.Y.Z`
   removes the entries. The gate's **branch coverage** check fails on a
   missing, stale or extra entry, as
   [docs/GOVERNANCE.md](docs/GOVERNANCE.md#dependabot) describes.
2. While the release branch is open, every change targets it and reaches
   `main` with the release, as [docs/GOVERNANCE.md](docs/GOVERNANCE.md#base-branch)
   describes. A CI, security or dependency fix that cannot wait goes to `main`
   directly; merge `main` into the release branch afterwards.
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
- `.github/dependabot.yml`: remove the `target-branch` entries of the release
  that `VERSION` now names
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

- Run `python3 tools/verify.py --ci` on the candidate. It runs every static
  check of the required **Repository and module checks** job, including the
  gate's checks of the 32-bit and 64-bit declaration arms and the house-style
  normal form. It needs Node.js, plus Ruff and mypy from
  `tools/requirements-quality-ci.txt`; with `--ci`, a check whose tool is
  missing fails instead of being skipped.
- Scan for secrets or generated files.

Capture commands, tool versions, timestamps, and complete results. Rerun affected gates after any change.

## 6. Certify in Excel

- Run **Debug → Compile VBAProject** on the exact candidate.
- Run `Test_EXCEL_UI_RunReleaseCertification`.
- Preferably run both with `tools\Certify-EXCEL_UI.cmd` from a clean checkout
  of the candidate, and validate the evidence folder it writes with
  `python3 tools/check_excel_evidence.py <folder> --commit <sha>`
  ([docs/EXCEL_CERTIFICATION.md](docs/EXCEL_CERTIFICATION.md)).
- Exercise hide/show, snapshot/restore, multi-window identity, and title-bar behavior.
- Confirm cleanup restores host UI state after success and error paths.

Certification rules:

- Use a clean workbook or documented clean fixture.
- Import only candidate files.
- Record Excel version, Windows version, and Office bitness.
- Certify on 64-bit Office, the only bitness the
  [support baseline](INSTALLATION.md) claims; 32-bit Office is not certified.
- Treat warnings, repairs, or unexplained numerical deltas as failures.
- If code changes after certification, restart static and Excel validation.

## 7. Build release artifacts

Planned outputs:

- Demo or certification `.xlsm` workbook, when published
- Source archives `VBA-EXCEL_UI-X.Y.Z.zip` and `VBA-EXCEL_UI-X.Y.Z.tar.gz`
  with `SHA256SUMS`, built from the tag by the Release workflow

The source archives need no manual step: the workflow builds them from the
pushed tag with `git archive`, so the `export-ignore` rules in `.gitattributes`
apply. A workbook is built by hand, as below.

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

Publish only artifacts promised by the installation guide. A workbook is not
part of the bundle, so it is not in `SHA256SUMS`: record its file name and
SHA-256 in the evidence record, attach it to the published release, and then
verify the release with it declared, as section 11 describes.

Disposable mutants and their local output belong only in `/.mutation-scratch/`.
They are not release evidence and must never enter source archives or binary
assets. Package from tracked candidate inputs, not a recursive working-tree
copy. The directory is ignored and marked `export-ignore` as defense in depth;
never force-add it. Reviewed control results may be retained separately with
their exact source identity, but disposable source variants are not deliverables.

## 8. Review and merge

Where policy requires a pull request, make these items easy to verify:

- intended version, previous tag, and candidate SHA;
- changelog and version synchronization;
- static and Excel results;
- artifact manifest and hashes;
- compatibility, migration, and security notes;
- remaining limitations.

Require configured checks and record the resulting `main` SHA. The merge must
leave `main` identical to the certified release branch:
`git diff --stat main release/vX.Y.Z` prints nothing. A difference means `main`
holds a change the release did not certify, or the release branch lacks a fix
made on `main`; resolve it before tagging. If the merge
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

Tag only the certified commit on `main`. A pushed `v*` tag cannot be deleted
or moved, and pushing it starts publication, so create the tag in your clone
and verify it there first:

```bash
git switch main
git pull --ff-only
git rev-parse HEAD
git tag -a vX.Y.Z -m "VBA Excel UI X.Y.Z"
git show --no-patch --decorate vX.Y.Z
git fetch origin main
python3 tools/release.py verify vX.Y.Z --repository danielep71/VBA-EXCEL_UI
```

`verify` reads the tag from your clone, so it runs before the tag is pushed. It
checks that the tag is annotated and equals `VERSION` at the tagged commit, that
`CHANGELOG.md` there has a dated section for the version, and that the commit is
reachable from `origin/main`, which is why `git fetch origin main` comes first.
It then asks GitHub for the **Repository and module checks** run on the tagged
commit's SHA. That commit is already on `main`, and every push to `main` runs
the check, so wait until that run has concluded.

Push the tag only when `verify` ends with `RESULT: PASS`:

```bash
git push origin vX.Y.Z
```

When it ends with `RESULT: FAIL`, do not push. If the check run on `main` was
only still in progress, run `verify` again once it has concluded. Otherwise
delete the local tag with `git tag -d vX.Y.Z` and fix the candidate; the new
`main` commit needs sections 5–8 again before you tag it.

The Release workflow runs the same verification again before it publishes
anything. Never delete and recreate a public tag to hide a mistake.

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

The Release workflow creates the release from the annotated tag. Its body is
the version's `CHANGELOG.md` section at the tagged commit, and its assets are
the source archives and `SHA256SUMS` from the bundle job; the publish job checks
the bundle against `SHA256SUMS` before it uploads anything. Write the release
text into that changelog section before tagging, including:

1. User-facing summary and highlights.
2. Upgrade or migration notes.
3. Supported platform statement.
4. Known limitations.
5. Installation link.
6. Full changelog comparison link.
7. Security-reporting link.

Do not edit the release body by hand: `verify-published` compares it with the
changelog section. Do not rebuild between tagging and upload.

## 11. Verify after publication

The publish job runs `verify-published`, which covers the release state, the
body and the asset hashes. The other checks below stay manual.

After attaching a workbook, run `verify-published` again with the workbook
declared by the name and SHA-256 in the evidence record. Download the bundle
artifact of the Release run first, because it holds the `SHA256SUMS` that the
workflow built; the run keeps it for 30 days:

```bash
gh run download <run-id> --repo danielep71/VBA-EXCEL_UI \
  --name release-bundle-<run-id> --dir dist
python3 tools/release.py verify-published vX.Y.Z \
  --repository danielep71/VBA-EXCEL_UI --dir dist \
  --asset <workbook>.xlsm=<sha256>
```

Repeat `--asset` for each additional asset. The check fails when an asset is
missing, when a hash differs, or when the release holds an asset that is neither
in `SHA256SUMS` nor declared.

| Check | Result |
| --- | :---: |
| Tag resolves to the certified SHA | ☐ |
| `VERSION` and changelog match the tag (`verify`) | ☐ |
| Assets download and hashes match (`verify-published`) | ☐ |
| Installation links and examples work | ☐ |
| Packaged artifact opens and passes its smoke test | ☐ |
| Source archive contains expected release files | ☐ |
| Default branch is ready for the next Unreleased cycle | ☐ |

Do not announce broad availability until these checks pass.

## 🧯 Recovery

### Before tag publication

Fix the branch or pull request, update evidence, and rerun affected gates. Rebuilding invalidates the prior hash and packaged-test result.

### After tag, before GitHub Release

If the verify job fails, nothing was published. The tag is protected and must
not be moved, so fix the cause in a new patch candidate. If the publish job
failed before **Create the GitHub release** finished, delete any draft release
it left behind and re-run the failed jobs from the workflow run.

If the release was created and only **Verify the published release** failed,
the release is already published, not a draft. Do not re-run the job: it
cannot create the release a second time. Download the bundle artifact and run
`verify-published` locally, as section 11 describes, and handle any finding
as **After public release** describes.

Pause publication and document the correction. Prefer a new patch candidate whenever the tag may have propagated.

### After public release

Do not silently replace assets or move the tag.

- Mark a dangerous release clearly and remove assets only for user safety.
- Publish a corrected patch with a new tag.
- Record the problem and remedy in the changelog.
- Use [SECURITY.md](SECURITY.md) for vulnerabilities.
- Preserve evidence explaining what users received.

## ☑️ Final release checklist

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
| 9 | Required pull request merged; `main` equals the release branch | ☐ |
| 10 | Annotated tag targets certified main SHA | ☐ |
| 11 | GitHub Release published by the Release workflow | ☐ |
| 12 | Post-publication checks pass | ☐ |

## 📚 Related documents

- [README.md](README.md) — project overview
- [INSTALLATION.md](INSTALLATION.md) — deployment and removal
- [CONTRIBUTING.md](CONTRIBUTING.md) — change and review workflow
- [docs/GOVERNANCE.md](docs/GOVERNANCE.md) — base branch, rulesets and repository settings
- [CHANGELOG.md](CHANGELOG.md) — release history
- [SECURITY.md](SECURITY.md) — supported versions and private reporting
- [VERSION](VERSION) — authoritative version
- [LICENSE](LICENSE) — MIT license

---

**Release principle:** certify one exact source revision, build from it once, and publish only evidence-backed artifacts derived from it.
