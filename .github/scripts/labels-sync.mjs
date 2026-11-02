#!/usr/bin/env node

import assert from "node:assert/strict";
import { appendFile, readFile } from "node:fs/promises";
import process from "node:process";
import { pathToFileURL } from "node:url";

const EXPECTED_SCHEMA_VERSION = 2;
const DEFAULT_MANIFEST = ".github/labels.json";
const API_VERSION = "2022-11-28";
const LABELS_PER_PAGE = 100;
const MAX_LABEL_PAGES = 100;
const MAX_RATE_LIMIT_RETRIES = 2;
export const MAX_RETRY_DELAY_MS = 5000;

export function compareNames(left, right) {
  const leftKey = left.toLowerCase();
  const rightKey = right.toLowerCase();
  if (leftKey < rightKey) return -1;
  if (leftKey > rightKey) return 1;
  return left < right ? -1 : left > right ? 1 : 0;
}

function sameKeys(value, expected) {
  const actual = Object.keys(value).sort();
  return JSON.stringify(actual) === JSON.stringify([...expected].sort());
}

export function normalizedLiveLabel(label) {
  return {
    name: String(label.name ?? ""),
    color: String(label.color ?? "").toUpperCase(),
    description: label.description == null ? "" : String(label.description)
  };
}

function validateLabelArray(labels, location, seen) {
  const errors = [];
  if (!Array.isArray(labels)) return [`${location} must be an array`];

  labels.forEach((label, index) => {
    const item = `${location}[${index}]`;
    if (!label || typeof label !== "object" || Array.isArray(label)) {
      errors.push(`${item} must be an object`);
      return;
    }
    if (!sameKeys(label, ["name", "color", "description"])) {
      errors.push(`${item} must contain exactly name, color, and description`);
    }

    if (typeof label.name !== "string") {
      errors.push(`${item}.name must be a string`);
    } else {
      if (label.name.length === 0 || label.name !== label.name.trim()) {
        errors.push(`${item}.name must be non-empty with no surrounding whitespace`);
      }
      if (label.name.length > 50) errors.push(`${item}.name exceeds GitHub's 50-character limit`);
      if (/\r|\n/.test(label.name)) errors.push(`${item}.name must be single-line`);
      const key = label.name.toLowerCase();
      if (seen.has(key)) {
        errors.push(`${item}.name duplicates ${seen.get(key)} case-insensitively`);
      } else {
        seen.set(key, `${item}.name`);
      }
    }

    if (typeof label.color !== "string" || !/^[0-9A-F]{6}$/.test(label.color)) {
      errors.push(`${item}.color must be six uppercase hexadecimal characters without #`);
    }
    if (typeof label.description !== "string") {
      errors.push(`${item}.description must be a string`);
    } else {
      if (label.description.length === 0) errors.push(`${item}.description must not be empty`);
      if (label.description.length > 100) errors.push(`${item}.description exceeds GitHub's 100-character limit`);
      if (/\r|\n/.test(label.description)) errors.push(`${item}.description must be single-line`);
    }
  });

  for (let index = 1; index < labels.length; index += 1) {
    const previous = labels[index - 1]?.name;
    const current = labels[index]?.name;
    if (typeof previous === "string" && typeof current === "string" && compareNames(previous, current) >= 0) {
      errors.push(`${location} must be sorted case-insensitively by name: ${previous} precedes ${current}`);
      break;
    }
  }
  return errors;
}

export function validateManifest(document) {
  const errors = [];
  if (!document || typeof document !== "object" || Array.isArray(document)) {
    throw new Error("Manifest must be a JSON object.");
  }
  if (!sameKeys(document, ["schema_version", "prune", "labels"])) {
    errors.push("root must contain exactly schema_version, prune, and labels");
  }
  if (document.schema_version !== EXPECTED_SCHEMA_VERSION) {
    errors.push(`schema_version must be ${EXPECTED_SCHEMA_VERSION}`);
  }
  if (typeof document.prune !== "boolean") errors.push("prune must be a boolean");

  errors.push(...validateLabelArray(document.labels, "labels", new Map()));
  if (Array.isArray(document.labels) && document.labels.length === 0) errors.push("labels must not be empty");

  if (errors.length > 0) throw new Error(`Invalid label manifest:\n- ${errors.join("\n- ")}`);
  return structuredClone(document);
}

export function resolveDesired(document) {
  const labels = [...validateManifest(document).labels];
  labels.sort((left, right) => compareNames(left.name, right.name));
  return labels;
}

export function planChanges(desiredLabels, liveLabels, { prune = true } = {}) {
  const desired = new Map(desiredLabels.map(label => [label.name.toLowerCase(), label]));
  const live = new Map(liveLabels.map(label => {
    const normalized = normalizedLiveLabel(label);
    return [normalized.name.toLowerCase(), normalized];
  }));
  const changes = [];

  for (const [key, target] of desired) {
    const current = live.get(key);
    if (!current) {
      changes.push({ action: "create", name: target.name, target });
      continue;
    }
    const fields = [];
    if (current.name !== target.name) fields.push("name");
    if (current.color !== target.color) fields.push("color");
    if (current.description !== target.description) fields.push("description");
    if (fields.length > 0) {
      changes.push({ action: "update", name: target.name, currentName: current.name, fields, target });
    }
  }

  if (prune) {
    for (const [key, current] of live) {
      if (!desired.has(key)) changes.push({ action: "delete", name: current.name, currentName: current.name });
    }
  }

  const order = { update: 0, create: 1, delete: 2 };
  changes.sort((left, right) => order[left.action] - order[right.action] || compareNames(left.name, right.name));
  return changes;
}

function parseArguments(argv) {
  const parsed = {
    manifest: DEFAULT_MANIFEST,
    mode: null,
    live: null,
    repository: process.env.GITHUB_REPOSITORY || null,
    dryRun: false,
    selfTest: false
  };
  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index];
    if (argument === "--dry-run") {
      parsed.dryRun = true;
    } else if (argument === "--self-test") {
      parsed.selfTest = true;
    } else if (["--manifest", "--mode", "--live", "--repository"].includes(argument)) {
      const value = argv[index + 1];
      if (!value || value.startsWith("--")) throw new Error(`${argument} requires a value`);
      parsed[argument.slice(2)] = value;
      index += 1;
    } else {
      throw new Error(`Unknown argument: ${argument}`);
    }
  }
  return parsed;
}

export async function loadJson(path) {
  let text;
  try {
    text = await readFile(path, "utf8");
  } catch (error) {
    throw new Error(`Cannot read ${path}: ${error.message}`);
  }
  try {
    return JSON.parse(text);
  } catch (error) {
    throw new Error(`Invalid JSON in ${path}: ${error.message}`);
  }
}

function splitRepository(repository) {
  const parts = String(repository ?? "").split("/");
  if (parts.length !== 2 || parts.some(part => part.length === 0)) {
    throw new Error("repository must be in owner/name form");
  }
  return parts.map(encodeURIComponent);
}

function apiUrl(repository, suffix) {
  const [owner, repo] = splitRepository(repository);
  // The runner sets GITHUB_API_URL, which differs from api.github.com on GitHub Enterprise.
  const base = (process.env.GITHUB_API_URL || "https://api.github.com").replace(/\/+$/, "");
  return `${base}/repos/${owner}/${repo}${suffix}`;
}

function sleep(milliseconds) {
  return new Promise(resolve => setTimeout(resolve, milliseconds));
}

function isRateLimited(response) {
  return response.status === 429
    || (
      response.status === 403
      && (
        response.headers.get("x-ratelimit-remaining") === "0"
        || response.headers.has("retry-after")
      )
    );
}

export function retryDelayMs(response, attempt) {
  const header = response.headers.get("retry-after");
  const retryAfter = header !== null && header.trim() !== "" ? Number(header) : NaN;
  if (Number.isFinite(retryAfter) && retryAfter >= 0) {
    return Math.min(retryAfter * 1000, MAX_RETRY_DELAY_MS);
  }
  return Math.min(1000 * (2 ** attempt), MAX_RETRY_DELAY_MS);
}

async function githubRequest(repository, token, method, suffix, body) {
  for (let attempt = 0; ; attempt += 1) {
    const response = await fetch(apiUrl(repository, suffix), {
      method,
      headers: {
        Accept: "application/vnd.github+json",
        Authorization: `Bearer ${token}`,
        "X-GitHub-Api-Version": API_VERSION,
        "User-Agent": "repository-labels-sync"
      },
      body: body === undefined ? undefined : JSON.stringify(body)
    });
    if (response.ok) {
      if (response.status === 204) return null;
      return response.json();
    }

    const detail = await response.text();
    const retryable = method === "GET" && isRateLimited(response);
    if (!retryable || attempt >= MAX_RATE_LIMIT_RETRIES) {
      throw new Error(`${method} ${suffix} failed with HTTP ${response.status}: ${detail.slice(0, 500)}`);
    }
    await sleep(retryDelayMs(response, attempt));
  }
}

export async function listLiveLabels(repository, token) {
  const labels = [];
  for (let page = 1; page <= MAX_LABEL_PAGES; page += 1) {
    const batch = await githubRequest(
      repository,
      token,
      "GET",
      `/labels?per_page=${LABELS_PER_PAGE}&page=${page}`
    );
    if (!Array.isArray(batch)) throw new Error("GitHub labels response is not an array");
    labels.push(...batch.map(normalizedLiveLabel));
    if (batch.length < LABELS_PER_PAGE) return labels;
  }
  throw new Error(`GitHub label pagination exceeded ${MAX_LABEL_PAGES} pages`);
}

async function applyChanges(repository, token, changes) {
  for (const change of changes) {
    if (change.action === "update") {
      await githubRequest(repository, token, "PATCH", `/labels/${encodeURIComponent(change.currentName)}`, {
        new_name: change.target.name,
        color: change.target.color,
        description: change.target.description
      });
    } else if (change.action === "create") {
      await githubRequest(repository, token, "POST", "/labels", change.target);
    } else if (change.action === "delete") {
      await githubRequest(repository, token, "DELETE", `/labels/${encodeURIComponent(change.currentName)}`);
    }
  }
}

function markdownEscape(value) {
  return String(value).replaceAll("|", "\\|").replaceAll("\n", " ");
}

function renderLabelNames(labels) {
  if (labels.length === 0) return "none";
  return labels.map(label => markdownEscape(JSON.stringify(label.name))).join(", ");
}

function renderSummary({
  mode,
  manifest,
  manifestPath,
  desiredLabels,
  changes = [],
  verified = null,
  dryRun = false,
  error = null,
  partial = false
}) {
  const resolvedLabels = desiredLabels ?? [];
  const planned = changes !== null;
  const counts = Object.fromEntries(["create", "update", "delete"].map(action => [
    action,
    planned ? changes.filter(change => change.action === action).length : "n/a"
  ]));
  const lines = [
    "# Repository label synchronization",
    "",
    `- Mode: \`${mode}\`${dryRun ? " (dry run)" : ""}`,
    `- Label manifest: \`${manifestPath ?? "unavailable"}\``,
    `- Schema: **${manifest?.schema_version ?? "unavailable"}**`,
    `- Labels: **${resolvedLabels.length}**`,
    `- Prune: **${manifest ? (manifest.prune ? "enabled" : "disabled") : "unavailable"}**`,
    `- Planned changes: **${planned ? changes.length : "not computed (live labels not read)"}**`,
    `- Create / update / delete: **${counts.create} / ${counts.update} / ${counts.delete}**`
  ];
  if (verified !== null) lines.push(`- Post-run exact match: **${verified ? "yes" : "no"}**`);
  lines.push(
    "",
    "## Label catalogue",
    "",
    `- Labels (${resolvedLabels.length}): ${renderLabelNames(resolvedLabels)}`
  );
  if (error) {
    lines.push("", "## Failure", "", `\`${markdownEscape(error)}\``);
    if (partial) {
      lines.push("", "The failure happened during or after reconciliation: some changes may already be applied. Compare the live labels with the manifest.");
    }
  } else if (!planned) {
    lines.push("", "Live labels were not read; no changes were planned.");
  } else if (changes.length === 0) {
    lines.push("", "No label changes are required.");
  } else {
    lines.push("", "| Action | Label | Detail |", "|---|---|---|");
    for (const change of changes) {
      const detail = change.action === "update" ? change.fields.join(", ") : change.action === "delete" ? "outside the manifest" : "missing live label";
      lines.push(`| ${change.action} | \`${markdownEscape(change.name)}\` | ${markdownEscape(detail)} |`);
    }
  }
  return `${lines.join("\n")}\n`;
}

async function publishSummary(summary) {
  process.stdout.write(summary);
  if (process.env.GITHUB_STEP_SUMMARY) await appendFile(process.env.GITHUB_STEP_SUMMARY, summary, "utf8");
}

function expectFailure(operation, pattern) {
  assert.throws(operation, pattern);
}

async function runSelfTest(manifestPath) {
  for (const [header, attempt, expected] of [
    [null, 0, 1000], [null, 2, 4000], ["", 1, 2000],
    ["invalid", 1, 2000], ["-1", 1, 2000], ["0", 2, 0],
    ["2", 0, 2000], ["999999", 0, MAX_RETRY_DELAY_MS],
    [null, 20, MAX_RETRY_DELAY_MS]
  ]) {
    const headers = new Headers(header === null ? {} : { "retry-after": header });
    assert.equal(retryDelayMs({ headers }, attempt), expected, `Retry-After: ${header}`);
  }

  const baseline = await loadJson(manifestPath);
  const manifest = validateManifest(baseline);
  const desired = resolveDesired(manifest);

  const missingRoot = structuredClone(baseline);
  delete missingRoot.prune;
  expectFailure(() => validateManifest(missingRoot), /root must contain exactly/);

  const extraRoot = structuredClone(baseline);
  extraRoot.overlays = {};
  expectFailure(() => validateManifest(extraRoot), /root must contain exactly/);

  const wrongSchema = structuredClone(baseline);
  wrongSchema.schema_version = 1;
  expectFailure(() => validateManifest(wrongSchema), /schema_version/);

  const invalidPrune = structuredClone(baseline);
  invalidPrune.prune = "true";
  expectFailure(() => validateManifest(invalidPrune), /prune must be a boolean/);

  const emptyLabels = structuredClone(baseline);
  emptyLabels.labels = [];
  expectFailure(() => validateManifest(emptyLabels), /must not be empty/);

  const missingField = structuredClone(baseline);
  delete missingField.labels[0].description;
  expectFailure(() => validateManifest(missingField), /name, color, and description/);

  const invalidColor = structuredClone(baseline);
  invalidColor.labels[0].color = "#c24e00";
  expectFailure(() => validateManifest(invalidColor), /uppercase hexadecimal/);

  const duplicateName = structuredClone(baseline);
  duplicateName.labels[1].name = baseline.labels[0].name.toUpperCase();
  expectFailure(() => validateManifest(duplicateName), /duplicates/);

  const longDescription = structuredClone(baseline);
  longDescription.labels[0].description = "x".repeat(101);
  expectFailure(() => validateManifest(longDescription), /100-character/);

  const unsorted = structuredClone(baseline);
  [unsorted.labels[0], unsorted.labels[1]] = [unsorted.labels[1], unsorted.labels[0]];
  expectFailure(() => validateManifest(unsorted), /sorted/);

  assert.equal(planChanges(desired, desired, { prune: manifest.prune }).length, 0, "idempotent plan must be empty");

  assert.deepEqual(planChanges(desired, desired.slice(1)).map(change => change.action), ["create"]);

  const changed = structuredClone(desired);
  changed[0].color = "000000";
  changed[0].description = "changed";
  const changedPlan = planChanges(desired, changed);
  assert.deepEqual(changedPlan.map(change => change.action), ["update"]);
  assert.deepEqual(changedPlan[0].fields, ["color", "description"]);

  const extra = [...structuredClone(desired), { name: "extra", color: "ABCDEF", description: "extra" }];
  assert.deepEqual(planChanges(desired, extra, { prune: true }).map(change => change.action), ["delete"]);
  assert.equal(planChanges(desired, extra, { prune: false }).length, 0);

  const caseDrift = structuredClone(desired);
  caseDrift[0].name = desired[0].name.toUpperCase();
  assert.deepEqual(planChanges(desired, caseDrift)[0].fields, ["name"]);

  const summary = renderSummary({
    mode: "plan",
    manifest,
    manifestPath: "fixture-labels.json",
    desiredLabels: desired,
    dryRun: true
  });
  assert.match(summary, /Label manifest: `fixture-labels\.json`/);
  assert.match(summary, new RegExp(`Labels: \\*\\*${desired.length}\\*\\*`));
  const validateSummary = renderSummary({
    mode: "validate",
    manifest,
    manifestPath: "fixture-labels.json",
    desiredLabels: desired,
    changes: null
  });
  assert.match(validateSummary, /Planned changes: \*\*not computed \(live labels not read\)\*\*/);
  assert.match(validateSummary, /Live labels were not read; no changes were planned\./);
  assert.doesNotMatch(validateSummary, /No label changes are required/);
  const failureSummary = renderSummary({
    mode: "failed",
    manifest: null,
    manifestPath: null,
    desiredLabels: [],
    changes: [{ action: "create", name: "new" }],
    partial: true,
    error: "write failed"
  });
  assert.match(failureSummary, /Prune: \*\*unavailable\*\*/);
  assert.match(failureSummary, /Planned changes: \*\*1\*\*/);
  assert.match(failureSummary, /some changes may already be applied/);
  assert.match(summary, new RegExp(`Labels \\(${desired.length}\\):`));
  for (const label of desired) assert.match(summary, new RegExp(`"${label.name}"`));

  process.stdout.write("SELF-TEST PASS: retry delays, validation and reconciliation fixtures\n");
}

// What the run knew before a failure, so that the failure summary reports it.
const runState = { summaryBase: null, changes: null, applying: false, verified: null };

async function main() {
  const options = parseArguments(process.argv.slice(2));
  if (options.selfTest) {
    await runSelfTest(options.manifest);
    return;
  }

  const manifest = validateManifest(await loadJson(options.manifest));
  const desired = resolveDesired(manifest);
  const summaryBase = {
    manifest,
    manifestPath: options.manifest,
    desiredLabels: desired
  };
  runState.summaryBase = summaryBase;

  if (options.mode === "validate") {
    await publishSummary(renderSummary({ mode: "validate", ...summaryBase, changes: null }));
    return;
  }

  let live;
  if (options.mode === "plan") {
    if (!options.live) throw new Error("--mode plan requires --live");
    const liveDocument = await loadJson(options.live);
    live = Array.isArray(liveDocument) ? liveDocument : liveDocument.labels;
    if (!Array.isArray(live)) throw new Error("live fixture must be an array or contain a labels array");
  } else if (options.mode === "reconcile") {
    if (!options.repository) throw new Error("--mode reconcile requires --repository or GITHUB_REPOSITORY");
    if (!process.env.GITHUB_TOKEN) throw new Error("--mode reconcile requires GITHUB_TOKEN");
    live = await listLiveLabels(options.repository, process.env.GITHUB_TOKEN);
  } else {
    throw new Error("--mode must be validate, plan, or reconcile");
  }

  const changes = planChanges(desired, live, { prune: manifest.prune });
  runState.changes = changes;
  if (options.mode === "plan" || options.dryRun) {
    await publishSummary(renderSummary({ mode: options.mode, ...summaryBase, changes, dryRun: options.dryRun }));
    return;
  }

  runState.applying = true;
  await applyChanges(options.repository, process.env.GITHUB_TOKEN, changes);
  const after = await listLiveLabels(options.repository, process.env.GITHUB_TOKEN);
  const remaining = planChanges(desired, after, { prune: manifest.prune });
  const verified = remaining.length === 0;
  if (!verified) {
    runState.verified = false;
    throw new Error(`Post-run verification found ${remaining.length} remaining difference(s)`);
  }
  await publishSummary(renderSummary({ mode: "reconcile", ...summaryBase, changes, verified }));
}

// Run only when executed directly, so that labels-drift.mjs can import the planner.
if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch(async error => {
    const message = error instanceof Error ? error.message : String(error);
    try {
      await publishSummary(renderSummary({
        mode: "failed",
        ...(runState.summaryBase ?? {
          manifest: null,
          manifestPath: null,
          desiredLabels: []
        }),
        changes: runState.changes,
        verified: runState.verified,
        partial: runState.applying,
        error: message
      }));
    } finally {
      process.stderr.write(`${message}\n`);
      process.exitCode = 1;
    }
  });
}
