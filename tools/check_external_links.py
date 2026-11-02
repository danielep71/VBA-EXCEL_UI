#!/usr/bin/env python3
"""Check the external links in the tracked Markdown files.

The static gate checks only links inside the repository, because a request to
another host can fail for reasons unrelated to the change under review. This
tool checks the external links instead. It runs on a weekly schedule, not as a
required check: a broken external link must be fixed, but it must not block a
merge.

Each unique URL is requested with bounded retries and timeouts:

- 2xx passes, after any redirects are followed;
- a redirect that never resolves (a loop, or no target) is broken;
- 404 and 410 are broken;
- 5xx, timeouts and connection errors are retried, then reported as broken;
- 401, 403 and 429 cannot be verified without credentials or a slower rate,
  so they are reported as unverified and do not fail the run.

URLs that match a prefix in tools/external_link_exclusions.txt are listed in
the report as excluded and not requested.

Exit status: 0 when no link is broken, 1 when at least one is, 2 on a usage or
repository error.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
import time
import urllib.error
import urllib.request
from dataclasses import asdict, dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
EXCLUSIONS = ROOT / "tools" / "external_link_exclusions.txt"
USER_AGENT = "VBA-EXCEL_UI-link-check (+https://github.com/danielep71/VBA-EXCEL_UI)"

# Inline links and images, autolinks, reference definitions, and the src or
# href of HTML img, a and source tags (a tag may span several lines). Code spans
# and fenced code are removed first, so example URLs in code are not requested.
INLINE = re.compile(r"\]\(\s*<?(https?://[^)\s>]+)>?(?:\s+\"[^\"]*\")?\s*\)")
AUTOLINK = re.compile(r"<(https?://[^>\s]+)>")
REFERENCE = re.compile(r"^\s{0,3}\[[^\]]+\]:\s*<?(https?://[^>\s]+)>?", re.M)
FENCE = re.compile(r"^(```|~~~).*?^\1[^\n]*$", re.M | re.S)
CODE_SPAN = re.compile(r"`[^`\n]*`")
HTML_ATTRIBUTE = re.compile(
    r"<(?:img|a|source)\b[^>]*?\s(?:src|href)\s*=\s*[\"'](https?://[^\"'\s>]+)[\"']",
    re.I,
)

BROKEN_STATUSES = {404, 410}
UNVERIFIED_STATUSES = {401, 403, 429}


@dataclass(frozen=True)
class Result:
    url: str
    outcome: str  # "ok", "broken", "unverified" or "excluded"
    detail: str
    sources: tuple[str, ...]


def tracked_markdown(root: Path) -> list[str]:
    output = subprocess.run(
        ["git", "ls-files", "-z", "--", "*.md"],
        cwd=root, capture_output=True, check=True,
    ).stdout.decode("utf-8")
    return sorted(path for path in output.split("\0") if path)


def extract_links(text: str) -> list[str]:
    """Return the external URLs of one Markdown document, in order of appearance."""
    text = FENCE.sub("", text)
    text = CODE_SPAN.sub("", text)
    found: list[tuple[int, str]] = []
    for pattern in (INLINE, AUTOLINK, REFERENCE, HTML_ATTRIBUTE):
        found += [(match.start(), match.group(1)) for match in pattern.finditer(text)]
    return [url.rstrip(".,;") for _, url in sorted(found)]


def load_exclusions(path: Path) -> list[str]:
    if not path.is_file():
        return []
    lines = path.read_text(encoding="utf-8").splitlines()
    return [line.strip() for line in lines if line.strip() and not line.lstrip().startswith("#")]


def collect(root: Path, files: list[str]) -> dict[str, list[str]]:
    """Map each unique URL, without its fragment, to the files that link to it."""
    sources: dict[str, list[str]] = {}
    for rel in files:
        text = (root / rel).read_text(encoding="utf-8", errors="replace")
        for url in extract_links(text):
            key = url.split("#", 1)[0]
            if rel not in sources.setdefault(key, []):
                sources[key].append(rel)
    return dict(sorted(sources.items()))


class UnresolvedRedirect(Exception):
    """A redirect chain ended without a final response (a loop or no target)."""


def request_status(url: str, timeout: float) -> int:
    """Return the final HTTP status of a GET request after following redirects.

    Raise UnresolvedRedirect when the redirects never reach a final response:
    urllib reports a redirect loop, or a redirect without a target, as an
    HTTPError carrying the 3xx status, which must not read as success.
    Raise on a transport error.
    """
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            status = int(response.status)
    except urllib.error.HTTPError as error:
        status = int(error.code)
    if 300 <= status < 400:
        raise UnresolvedRedirect(f"HTTP {status}: the redirect did not resolve")
    return status


def check_url(url: str, attempts: int, timeout: float, delay: float) -> tuple[str, str]:
    detail = ""
    for attempt in range(1, attempts + 1):
        try:
            status = request_status(url, timeout)
        except UnresolvedRedirect as error:
            return "broken", str(error)
        except (urllib.error.URLError, TimeoutError, OSError) as error:
            detail = f"transport error: {error}"
        else:
            if status < 300:
                return "ok", f"HTTP {status}"
            if status in BROKEN_STATUSES:
                return "broken", f"HTTP {status}"
            if status in UNVERIFIED_STATUSES:
                return "unverified", f"HTTP {status}"
            detail = f"HTTP {status}"
        if attempt < attempts:
            time.sleep(delay * attempt)
    return "broken", f"{detail} after {attempts} attempts"


def render(results: list[Result]) -> str:
    counts = {name: sum(1 for r in results if r.outcome == name)
              for name in ("ok", "broken", "unverified", "excluded")}
    status = "FAIL" if counts["broken"] else "PASS"
    lines = [
        "# External link check",
        "",
        f"Status: **{status}** · checked {len(results)} unique URLs · "
        + " · ".join(f"{name}: {count}" for name, count in counts.items()),
        "",
    ]
    attention = [r for r in results if r.outcome in ("broken", "unverified")]
    if attention:
        lines += ["| Outcome | URL | Detail | Linked from |", "| --- | --- | --- | --- |"]
        lines += [f"| {r.outcome} | {r.url} | {r.detail} | {', '.join(r.sources)} |"
                  for r in attention]
    else:
        lines.append("Every checked link resolved.")
    return "\n".join(lines) + "\n"


def run(root: Path, attempts: int, timeout: float, delay: float) -> list[Result]:
    exclusions = load_exclusions(root / EXCLUSIONS.relative_to(ROOT))
    results = []
    for url, files in collect(root, tracked_markdown(root)).items():
        if any(url.startswith(prefix) for prefix in exclusions):
            outcome, detail = "excluded", "listed in tools/external_link_exclusions.txt"
        else:
            outcome, detail = check_url(url, attempts, timeout, delay)
        results.append(Result(url, outcome, detail, tuple(files)))
    return results


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0] if __doc__ else None)
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--attempts", type=int, default=3)
    parser.add_argument("--timeout", type=float, default=20.0)
    parser.add_argument("--delay", type=float, default=5.0,
                        help="seconds before the first retry; later retries wait longer")
    parser.add_argument("--output", type=Path, help="write the results as JSON")
    parser.add_argument("--summary", type=Path, help="append the Markdown report to this file")
    parser.add_argument("--list", action="store_true",
                        help="print the URLs and their sources without requesting them")
    args = parser.parse_args(argv)
    if args.attempts < 1 or args.timeout <= 0 or args.delay < 0:
        parser.error("--attempts must be at least 1, --timeout positive, --delay not negative")

    root = args.root.resolve()
    try:
        if args.list:
            for url, files in collect(root, tracked_markdown(root)).items():
                print(f"{url}\t{', '.join(files)}")
            return 0
        results = run(root, args.attempts, args.timeout, args.delay)
    except (OSError, subprocess.CalledProcessError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2

    report = render(results)
    print(report, end="")
    if args.summary:
        with args.summary.open("a", encoding="utf-8") as handle:
            handle.write(report)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps([asdict(r) for r in results], indent=2) + "\n",
                               encoding="utf-8")
    return 1 if any(r.outcome == "broken" for r in results) else 0


if __name__ == "__main__":
    raise SystemExit(main())
