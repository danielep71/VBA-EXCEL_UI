#!/usr/bin/env node

import assert from "node:assert/strict";
import { appendFile, mkdir, writeFile } from "node:fs/promises";
import process from "node:process";

// The synchronization script owns validation, planning and the GitHub client;
// importing them keeps the drift report and the reconciler on one plan.
import {
  MAX_RETRY_DELAY_MS,
  compareNames,
  listLiveLabels,
  loadJson,
  normalizedLiveLabel,
  planChanges,
  resolveDesired,
  retryDelayMs,
  validateManifest
} from "./labels-sync.mjs";

const TOOL_NAME = "Repository label drift";
const DEFAULT_MANIFEST = ".github/labels.json";

function parseArguments(argv) {
  const parsed = {
    manifest: DEFAULT_MANIFEST,
    live: null,
    repository: process.env.GITHUB_REPOSITORY || null,
    output: null,
    summary: null,
    selfTest: false
  };
  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index];
    if (argument === "--self-test") {
      parsed.selfTest = true;
      continue;
    }
    if (["--manifest", "--live", "--repository", "--output", "--summary"].includes(argument)) {
      const value = argv[index + 1];
      if (!value || value.startsWith("--")) throw new Error(`${argument} requires a value`);
      parsed[argument.slice(2)] = value;
      index += 1;
      continue;
    }
    throw new Error(`Unknown argument: ${argument}`);
  }
  return parsed;
}

function changeEvidence(change, observedByName) {
  const currentKey = String(change.currentName ?? change.name).toLowerCase();
  return {
    action: change.action,
    name: change.name,
    fields: change.fields ?? [],
    observed: observedByName.get(currentKey) ?? null,
    desired: change.target ?? null
  };
}

function buildReport({ repository, manifestPath, manifest, desiredLabels, liveLabels }) {
  const observed = liveLabels
    .map(normalizedLiveLabel)
    .sort((left, right) => compareNames(left.name, right.name));
  const desired = structuredClone(desiredLabels)
    .sort((left, right) => compareNames(left.name, right.name));
  const changes = planChanges(desired, observed, { prune: manifest.prune });
  const observedByName = new Map(observed.map(label => [label.name.toLowerCase(), label]));
  const differences = changes.map(change => changeEvidence(change, observedByName));
  const counts = Object.fromEntries(["create", "update", "delete"].map(action => [
    action,
    differences.filter(item => item.action === action).length
  ]));
  return {
    schema_version: 2,
    tool: TOOL_NAME,
    status: differences.length === 0 ? "pass" : "drift",
    repository,
    manifest: manifestPath,
    prune: manifest.prune,
    counts: {
      desired: desired.length,
      observed: observed.length,
      differences: differences.length,
      create: counts.create,
      update: counts.update,
      delete: counts.delete
    },
    desired,
    observed,
    differences,
    // The plan comes from the synchronization script's own planner.
    canonical_plan_verified: true
  };
}

function markdownEscape(value) {
  return String(value).replaceAll("|", "\\|").replaceAll("\n", " ");
}

function compactLabel(label) {
  if (label === null) return "—";
  return `\`${markdownEscape(label.name)}\` / \`${label.color}\` / ${markdownEscape(label.description)}`;
}

function markdownReport(report) {
  const lines = [
    "# Repository label drift",
    "",
    `- Status: **${String(report.status).toUpperCase()}**`,
    `- Repository: \`${report.repository ?? "fixture"}\``,
    `- Manifest: \`${report.manifest}\``,
    `- Desired / observed: **${report.counts.desired} / ${report.counts.observed}**`,
    `- Differences: **${report.counts.differences}**`,
    `- Create / update / delete: **${report.counts.create} / ${report.counts.update} / ${report.counts.delete}**`,
    "- Plan: computed by the synchronization script's planner",
    ""
  ];
  if (report.differences.length === 0) {
    lines.push("No live label drift detected.");
  } else {
    lines.push(
      "| Action | Label | Fields | Observed | Desired |",
      "| --- | --- | --- | --- | --- |"
    );
    for (const difference of report.differences) {
      lines.push(
        `| ${difference.action} | \`${markdownEscape(difference.name)}\` | `
        + `${difference.fields.length > 0 ? difference.fields.join(", ") : "—"} | `
        + `${compactLabel(difference.observed)} | ${compactLabel(difference.desired)} |`
      );
    }
  }
  return `${lines.join("\n")}\n`;
}

async function writeEvidence(path, content) {
  if (!path) return;
  const slash = path.lastIndexOf("/");
  if (slash > 0) await mkdir(path.slice(0, slash), { recursive: true });
  await writeFile(path, content, "utf8");
}

async function publishSummary(summary) {
  process.stdout.write(summary);
  if (process.env.GITHUB_STEP_SUMMARY) {
    await appendFile(process.env.GITHUB_STEP_SUMMARY, summary, "utf8");
  }
}

async function prepareContract(manifestPath) {
  const manifest = validateManifest(await loadJson(manifestPath));
  return { manifest, desiredLabels: resolveDesired(manifest) };
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

  const { manifest, desiredLabels } = await prepareContract(manifestPath);

  const baseline = structuredClone(desiredLabels);
  const baselineBefore = JSON.stringify(baseline);
  const cleanFirst = buildReport({
    repository: "fixture/clean",
    manifestPath,
    manifest,
    desiredLabels,
    liveLabels: baseline
  });
  const cleanSecond = buildReport({
    repository: "fixture/clean",
    manifestPath,
    manifest,
    desiredLabels,
    liveLabels: baseline
  });
  assert.equal(cleanFirst.status, "pass");
  assert.equal(cleanFirst.counts.differences, 0);
  assert.equal(JSON.stringify(cleanFirst), JSON.stringify(cleanSecond));
  assert.equal(markdownReport(cleanFirst), markdownReport(cleanSecond));
  assert.equal(JSON.stringify(baseline), baselineBefore, "clean check must not mutate live input");

  const drifted = structuredClone(desiredLabels);
  const missing = drifted.shift();
  assert.ok(missing, "fixture requires at least one desired label");
  drifted[0].color = drifted[0].color === "000000" ? "FFFFFF" : "000000";
  drifted[0].description = "simulated out-of-band change";
  drifted.push({ name: "out-of-band", color: "ABCDEF", description: "simulated extra label" });
  const driftedBefore = JSON.stringify(drifted);
  const driftFirst = buildReport({
    repository: "fixture/drift",
    manifestPath,
    manifest,
    desiredLabels,
    liveLabels: drifted
  });
  const driftSecond = buildReport({
    repository: "fixture/drift",
    manifestPath,
    manifest,
    desiredLabels,
    liveLabels: drifted
  });
  assert.equal(driftFirst.status, "drift");
  assert.deepEqual(driftFirst.differences.map(item => item.action), ["update", "create", "delete"]);
  assert.equal(JSON.stringify(driftFirst), JSON.stringify(driftSecond));
  assert.equal(markdownReport(driftFirst), markdownReport(driftSecond));
  assert.equal(JSON.stringify(drifted), driftedBefore, "drift check must not mutate live input");
  assert.match(markdownReport(driftFirst), /simulated out-of-band change/);
  assert.match(markdownReport(driftFirst), new RegExp(missing.name));

  // Names with Markdown syntax are data, not table syntax.
  const special = [...structuredClone(desiredLabels),
    { name: "needs|triage", color: "ABCDEF", description: "pipe" },
    { name: "wip `x`", color: "ABCDEF", description: "backtick" }];
  const specialReport = buildReport({
    repository: "fixture/special", manifestPath, manifest, desiredLabels, liveLabels: special
  });
  assert.equal(specialReport.status, "drift");
  assert.deepEqual(specialReport.differences.map(item => [item.action, item.name]),
    [["delete", "needs|triage"], ["delete", "wip `x`"]]);
  assert.match(markdownReport(specialReport), /needs\\\|triage/);

  process.stdout.write(
    "SELF-TEST PASS: retry delays and deterministic no-drift, create/update/delete and special-name drift fixtures are read-only.\n"
  );
}

async function main() {
  const options = parseArguments(process.argv.slice(2));
  if (options.selfTest) {
    await runSelfTest(options.manifest);
    return;
  }

  const { manifest, desiredLabels } = await prepareContract(options.manifest);

  let liveLabels;
  if (options.live) {
    const liveDocument = await loadJson(options.live);
    liveLabels = Array.isArray(liveDocument) ? liveDocument : liveDocument.labels;
    if (!Array.isArray(liveLabels)) {
      throw new Error("live fixture must be an array or contain a labels array");
    }
  } else {
    if (!options.repository) throw new Error("live check requires --repository or GITHUB_REPOSITORY");
    if (!process.env.GITHUB_TOKEN) throw new Error("live check requires GITHUB_TOKEN");
    liveLabels = await listLiveLabels(options.repository, process.env.GITHUB_TOKEN);
  }

  const report = buildReport({
    repository: options.repository,
    manifestPath: options.manifest,
    manifest,
    desiredLabels,
    liveLabels
  });
  const json = `${JSON.stringify(report, null, 2)}\n`;
  const markdown = markdownReport(report);
  await writeEvidence(options.output, json);
  await writeEvidence(options.summary, markdown);
  await publishSummary(markdown);
  process.exitCode = report.status === "pass" ? 0 : 1;
}

main().catch(error => {
  const message = error instanceof Error ? error.message : String(error);
  process.stderr.write(`ERROR: ${message}\n`);
  process.exitCode = 2;
});
