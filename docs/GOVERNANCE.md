# 🏛️ Repository governance

These are the rules that protect the repository's branches and history: which
branch a change targets, the rulesets and settings GitHub enforces, how pull
requests and review threads are merged, how pinned dependencies are updated,
and where the issue-label rules live.
[CONTRIBUTING.md](../CONTRIBUTING.md) describes how to prepare and submit a
change within them.

<a id="base-branch"></a>

## Base branch

While no release branch is open, start from the current `main`. While a
release branch `release/vX.Y.Z` is open, every change targets it: VBA
corrections, tooling, workflows and documentation alike. `main` receives that
work in one merge, once the release is certified, so I never copy a change to
both branches.

While a release branch is open, a change goes to `main` directly only when it
fixes CI, security or a dependency and cannot wait for the release, such as a
Dependabot update of a pinned Action. After such a change, merge `main` into
the release branch in a pull request, so the release still contains everything
on `main`.

<a id="rulesets"></a>

## Rulesets, pull requests and review threads

Every change to `main` and `release/**` goes through a pull request, including
mine. Repository rulesets enforce it: on `main` and `release/**` a pull request
must pass **Repository and module checks** on an up-to-date branch with every
review thread resolved, and is merged with a merge commit; a push to the pull
request dismisses earlier approvals. Force-pushes and deletion are blocked, and
`v*` tags cannot be moved or deleted. The rulesets are versioned in `.github/rulesets/`,
and the **Ruleset drift** workflow fails daily if a live ruleset differs from
its copy. To change a ruleset, change it in the repository settings, run
`python3 tools/check_rulesets.py --repository danielep71/VBA-EXCEL_UI --export`,
and commit the result in a pull request that explains why.

<a id="repository-settings"></a>

## Repository settings

The other settings that protect the repository are versioned in
`.github/repository-settings.json`: the merge methods and repository features,
secret scanning, push protection, Dependabot alerts and security updates,
private vulnerability reporting, and the GitHub Actions policy (allowed actions,
the default `GITHUB_TOKEN` permissions, and whether Actions may approve pull
requests). The same **Ruleset drift** workflow runs `tools/check_settings.py`,
which reports each setting as a match, a difference or not verifiable, and
fails only on a difference. Its token is read-only, so the security and Actions
settings, which GitHub shows only to administrators, are reported as not
verifiable there. I check every setting locally with my own administrator
token:

```text
GITHUB_TOKEN="$(gh auth token)" python3 tools/check_settings.py --repository danielep71/VBA-EXCEL_UI --strict
```

`--strict` also fails when a setting is not verifiable, so a token without
administrator access cannot pass. To change a setting, change it in the
repository settings, update `.github/repository-settings.json` in a pull
request that explains why, and run the command again.

## Updating a pinned Action

Every `uses:` reference in `.github/workflows/` is pinned to a full 40-character
commit SHA with a trailing comment naming the release it is, and
`tools/check_repo.py` fails on anything else. A version tag is a mutable
pointer: `actions/checkout@v4` moved to a different commit as recently as the
pin currently in the tree, so a workflow referencing it runs whatever the
upstream account publishes next, including whatever an attacker publishes after
compromising it.

The comment is not decoration. A bare 40-hex string cannot be reviewed — nobody
can tell a real release from an arbitrary commit by reading it. The comment
states the claim, and resolving the tag is what checks it.

To move a pin:

1. Find the intended release in the official Action repository and note its tag,
   for example `v4.4.0`. Read its release notes; a pin bump is a dependency
   change, not housekeeping.
2. Resolve that tag to a commit without trusting a web page:

   ```text
   git ls-remote https://github.com/actions/checkout.git refs/tags/v4.4.0
   ```

   If the tag is annotated, resolve `refs/tags/v4.4.0^{}` as well and pin the
   commit it dereferences to, not the tag object.
3. Replace the SHA and the version comment **together**, in one edit. A SHA
   updated without its comment, or a comment updated without its SHA, is worse
   than no pin: it is a false statement a reviewer will believe.
4. Run `python3 tools/check_repo.py`.

A repository-local action — `uses: ./…` — is versioned by this repository's own
history and is exempt. Nothing else is.

## Dependabot

Dependabot (`.github/dependabot.yml`) opens a weekly grouped pull request when a
pinned Action has a newer release, and a monthly one for the hash-locked Python
quality tools. It changes the SHA and the version comment together; review it
with the steps in [Updating a pinned Action](#updating-a-pinned-action) before
merging. While a release branch is open, Dependabot
opens the same updates against it too, so neither branch falls behind; remove
those entries when the release branch is merged.

The gate's **branch coverage** check holds that configuration to the branches
that exist. Every ecosystem needs one entry without `target-branch`, for
`main`. While a release is open, it also needs one entry for that release
branch, with the same directory, schedule, groups and labels as the `main`
entry. The release branch must be the next patch, minor or major release after
`VERSION`; an entry for a release that has already shipped, or for any other
branch, fails. The same check requires **CodeQL** to run on pushes to `main`
and on pull requests to `main` and `release/**`, and rejects negative (`!`)
patterns in those branch lists.

The actionlint release that **Workflow lint** downloads is neither an Action nor
a Python package, so Dependabot does not update it. Update
`ACTIONLINT_VERSION` and `ACTIONLINT_SHA256` in
`.github/workflows/actionlint.yml` by hand, together, in one edit.

## Labels and milestones

[ISSUE_LABELS.md](ISSUE_LABELS.md) is the reference for issue labels,
priorities, milestones, target dates, dependency status and the workflows
that keep them consistent. `.github/labels.json` is the authoritative label
catalogue; change it through a pull request.
