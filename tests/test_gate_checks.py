"""One passing and at least one failing fixture for every check in CHECKS.

A gate check that silently stops detecting its defect still prints ``ok``.
Each class below runs one check against a copy of this repository: unchanged,
the check must report nothing; with one deliberate defect, it must report a
finding that names it. The self-test checks are proved the same way, by
breaking the rule their fixtures exercise and expecting the self-test to say so.
"""

import hashlib
import json
import re
import unittest
from collections.abc import Callable
from typing import Any
from unittest.mock import patch

from gate_fixture import GateTestCase, check_repo, reformat, vba_api, wiki_badges

RUNTIME = "src/M_EXCEL_UI_RUNTIME.bas"
TESTS = "test/M_EXCEL_UI_REGRESSION_TESTS.bas"
BUILDER = "demo/M_DEMO_BUILDER.bas"
WORKFLOW = ".github/workflows/static-checks.yml"
RULE = "'" + "=" * 78 + "\r\n"
NEVER = re.compile(r"(?!)")


class RepositoryPolicy(GateTestCase):
    check = "repository policy"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_when_a_broad_rule_ignores_authoritative_source(self) -> None:
        self.repo.append(".gitignore", "\n*.bas\n")
        self.assertFails("Authoritative path ignored: src/mutant.bas")


class RepositoryPolicySelfTest(GateTestCase):
    check = "repository policy self-test"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_when_the_baseline_policy_is_already_broken(self) -> None:
        self.repo.replace(".gitignore", "/.mutation-scratch/\n", "")
        self.assertFails("baseline")


class ModuleInventory(GateTestCase):
    check = "module inventory"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_on_a_module_outside_the_module_directories(self) -> None:
        self.repo.write("docs/M_STRAY.bas", 'Attribute VB_Name = "M_STRAY"\r\n')
        self.assertFails("docs/M_STRAY.bas: VBA module tracked outside")

    def test_fails_on_a_module_in_a_nested_directory(self) -> None:
        self.repo.write("test/fixtures/M_NESTED.bas", 'Attribute VB_Name = "M_NESTED"\r\n')
        self.assertFails("test/fixtures/M_NESTED.bas: VBA module tracked outside")

    def test_fails_on_an_undocumented_production_module(self) -> None:
        self.repo.write("src/M_EXCEL_UI_EXTRA.bas", 'Attribute VB_Name = "M_EXCEL_UI_EXTRA"\r\n')
        self.assertFails("src/M_EXCEL_UI_EXTRA.bas: production module is not in EXPECTED_PRODUCTION_MODULES",
                         "src/M_EXCEL_UI_EXTRA.bas: production module in no public API section")

    def test_fails_on_a_documented_production_module_that_is_not_tracked(self) -> None:
        self.repo.remove("src/M_EXCEL_UI_SNAPSHOT.bas")
        self.assertFails("src/M_EXCEL_UI_SNAPSHOT.bas: listed in EXPECTED_PRODUCTION_MODULES but not tracked")


class VbaProjectManifest(GateTestCase):
    check = "VBA project manifest"
    MANIFEST = "tools/vba-projects.json"

    def edit(self, change: Callable[[dict[str, Any]], object]) -> None:
        manifest = json.loads(self.repo.read_text(self.MANIFEST))
        change(manifest)
        self.repo.write(self.MANIFEST, json.dumps(manifest, indent=2) + "\n")

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_when_a_module_is_missing_from_the_manifest(self) -> None:
        self.edit(lambda m: m["configurations"]["development"]["modules"].remove(TESTS))
        self.assertFails(f"development does not list {TESTS}",
                         f"{TESTS}: tracked module is in no configuration of {self.MANIFEST}")

    def test_fails_on_a_new_module_that_no_configuration_lists(self) -> None:
        self.repo.write("demo/M_DEMO_EXTRA.bas", 'Attribute VB_Name = "M_DEMO_EXTRA"\r\n')
        self.assertFails("demo does not list demo/M_DEMO_EXTRA.bas",
                         "demo/M_DEMO_EXTRA.bas: tracked module is in no configuration")

    def test_fails_on_a_listed_path_that_does_not_exist(self) -> None:
        self.edit(lambda m: m["configurations"]["production"]["modules"].append("src/M_GONE.bas"))
        self.assertFails("production lists src/M_GONE.bas, which does not exist")

    def test_fails_on_a_listed_file_that_git_does_not_track(self) -> None:
        self.repo.write("src/M_LOCAL.bas", 'Attribute VB_Name = "M_LOCAL"\r\n', track=False)
        self.edit(lambda m: m["configurations"]["production"]["modules"].append("src/M_LOCAL.bas"))
        self.assertFails("production lists src/M_LOCAL.bas, which Git does not track")

    def test_fails_on_a_module_in_the_wrong_configuration_or_listed_twice(self) -> None:
        def change(manifest: dict[str, Any]) -> None:
            manifest["configurations"]["production"]["modules"].append(TESTS)
            manifest["configurations"]["demo"]["modules"].append(BUILDER)
        self.edit(change)
        self.assertFails(f"production lists {TESTS}, which does not belong to it",
                         f"demo lists {BUILDER} more than once")

    def test_fails_on_document_modules_or_a_missing_configuration(self) -> None:
        def change(manifest: dict[str, Any]) -> None:
            manifest["document_modules"] = ["src/ThisWorkbook.cls"]
            del manifest["configurations"]["demo"]
        self.edit(change)
        self.assertFails("document_modules must be an empty list",
                         "configuration 'demo' has no list of module paths",
                         f"{BUILDER}: tracked module is in no configuration")

    def test_fails_on_an_invalid_pattern(self) -> None:
        self.edit(lambda m: m["project_name_patterns"].append("^M_("))
        self.assertFails("project name pattern '^M_(' is invalid")

    def test_fails_on_a_manifest_that_does_not_parse_or_is_missing(self) -> None:
        self.repo.write(self.MANIFEST, "{ not json\n")
        self.assertFails(f"{self.MANIFEST} is not valid JSON")
        self.repo.remove(self.MANIFEST)
        self.assertFails(f"{self.MANIFEST} is missing")


class RequiredFiles(GateTestCase):
    check = "required files"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_on_a_missing_file(self) -> None:
        self.repo.remove("LICENSE")
        self.assertFails("missing LICENSE")

    def test_fails_on_a_directory_in_place_of_a_file(self) -> None:
        self.repo.remove("VERSION")
        self.repo.path("VERSION").mkdir()
        self.assertFails("VERSION is a directory, not a file")


class ModuleNames(GateTestCase):
    check = "module names"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_on_a_name_that_differs_from_the_file(self) -> None:
        self.repo.replace(RUNTIME, '"M_EXCEL_UI_RUNTIME"', '"M_WRONG"')
        self.assertFails(f"{RUNTIME}: declares 'M_WRONG', expected 'M_EXCEL_UI_RUNTIME'")

    def test_fails_when_line_one_is_not_the_name_attribute(self) -> None:
        self.repo.write(RUNTIME, b"' comment\r\n" + self.repo.read_bytes(RUNTIME))
        self.assertFails(f"{RUNTIME}: no Attribute VB_Name on line 1")


class OptionPolicy(GateTestCase):
    check = "option policy"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_without_option_explicit(self) -> None:
        self.repo.replace(BUILDER, "Option Explicit\r\n", "")
        self.assertFails(f"{BUILDER}: missing Option Explicit")

    def test_fails_without_option_private_module_in_a_test_module(self) -> None:
        self.repo.replace(TESTS, "Option Private Module\r\n", "")
        self.assertFails(f"{TESTS}: missing Option Private Module")


class EncodingAndLineEndings(GateTestCase):
    check = "encoding and line endings"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_on_each_defect_class(self) -> None:
        self.repo.append(RUNTIME, b"x \t\xe9\n")
        self.assertFails(f"{RUNTIME}: 1 bare LF", f"{RUNTIME}: contains tab characters",
                         f"{RUNTIME}: 1 non-ASCII bytes", "trailing whitespace")


class BannerRuleWidths(GateTestCase):
    check = "banner rule widths"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_on_a_short_rule(self) -> None:
        self.repo.replace(RUNTIME, RULE, "'" + "=" * 70 + "\r\n")
        self.assertFails("rule is 71 cols, expected 79")


class HouseLabels(GateTestCase):
    check = "house labels"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_on_a_label_outside_the_house_vocabulary(self) -> None:
        self.repo.replace(RUNTIME, "\r\nEnd Sub\r\n", "\r\nBogus_Label:\r\nEnd Sub\r\n")
        self.assertFails("non-house label 'Bogus_Label'")


class VbaAnalyzer(GateTestCase):
    check = "VBA analyzer"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_on_an_unresolved_project_reference(self) -> None:
        self.repo.replace(RUNTIME, "\r\nEnd Sub\r\n", "\r\n    UI_DoesNotExist\r\nEnd Sub\r\n")
        self.assertFails("unresolved project reference UI_DoesNotExist")


class VbaAnalyzerSelfTest(GateTestCase):
    check = "VBA analyzer self-test"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_when_the_analyzer_reports_nothing(self) -> None:
        with patch("vba_analyze_fixtures.analyze", return_value=([], [])):
            self.assertFails("unresolved direct call")


class TitleBarControlRegression(GateTestCase):
    check = "title-bar control regression"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_when_cleanup_leaves_its_handler_armed(self) -> None:
        text = self.repo.read_bytes(TESTS).decode("ascii")
        start = text.index("Sub TST_Case_TitleBarShowRejectsCaptionlessBaseline")
        cleanup = text.index("Safe_Exit:", start)
        disarm = text.index("On Error GoTo 0", cleanup)
        text = text[:disarm] + "On Error Resume Next" + text[disarm + len("On Error GoTo 0"):]
        self.repo.write(TESTS, text.encode("ascii"))
        self.assertFails("cleanup re-raise must disable its handler")

    def test_fails_when_a_restore_result_is_discarded(self) -> None:
        text = self.repo.read_bytes(TESTS).decode("ascii")
        start = text.index("Sub TST_Case_TitleBarShowRejectsCaptionlessBaseline")
        checked = "        RestoreOK = TST_RestoreTitleBarStyle( _\r\n"
        call = text.index(checked, start)
        text = text[:call] + "        TST_RestoreTitleBarStyle _\r\n" + text[call + len(checked):]
        self.repo.write(TESTS, text.encode("ascii"))
        self.assertFails("TST_RestoreTitleBarStyle result discarded; check it and report a failed restoration")


class WinApiErrorCapture(GateTestCase):
    check = "WinAPI error capture"
    TITLEBAR = "src/M_EXCEL_UI_TITLEBAR.bas"
    CAPTURE = "        DllError = Err.LastDllError\r\n"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_when_a_call_runs_between_the_api_and_the_capture(self) -> None:
        text = self.repo.read_bytes(self.TITLEBAR).decode("ascii")
        start = text.index("Private Function UI_TrySetWindowStyle(")
        capture = text.index(self.CAPTURE, start)
        text = (text[:capture] + "        FailMsg = UI_TitleBarBuildRuntimeErrorText\r\n"
                + text[capture:])
        self.repo.write(self.TITLEBAR, text.encode("ascii"))
        self.assertFails("UI_TrySetWindowStyle: Err.LastDllError read after an intervening statement",
                         "SetWindowLongPtr is not immediately followed by its Err.LastDllError capture")

    def test_fails_when_on_error_runs_between_the_api_and_the_capture(self) -> None:
        text = self.repo.read_bytes(TESTS).decode("ascii")
        start = text.index("Private Function TST_TryRefreshWindowFrame(")
        capture = text.index(self.CAPTURE, start)
        text = text[:capture] + "        On Error Resume Next\r\n" + text[capture:]
        self.repo.write(TESTS, text.encode("ascii"))
        self.assertFails("TST_TryRefreshWindowFrame: Err.LastDllError read after an On Error statement")

    def test_fails_when_only_the_legacy_arm_loses_its_clear(self) -> None:
        call = "        ApiResult = GetWindowLong(hWnd, GWL_STYLE)\r\n"
        self.repo.replace(self.TITLEBAR, f"\r\n#Else\r\n{call}#End If\r\n",
                          f"\r\n#Else\r\n        FaultError = 0\r\n{call}#End If\r\n")
        found = self.assertFails("legacy-x86: UI_TryGetWindowStyle: GetWindowLong is not immediately "
                                 "preceded by SetLastError 0")
        self.assertFalse([f for f in found if "vba7-" in f], found)

    def test_fails_when_get_last_error_is_declared_again(self) -> None:
        clear = '        Private Declare Sub TST_SetLastError Lib "kernel32"'
        declare = '        Private Declare Function TST_GetLastError Lib "kernel32" Alias "GetLastError" () As Long'
        self.repo.replace(TESTS, clear, declare + "\r\n" + clear)
        self.assertFails("legacy-x86: declares TST_GetLastError; capture Err.LastDllError instead")

    def test_fails_when_the_fixtures_stop_detecting_a_late_capture(self) -> None:
        with patch("winapi_capture_fixtures.procedure_findings", return_value=iter(())):
            self.assertFails("#27 fixture a helper call between the call and the capture is rejected")


class DemoCleanupRegression(GateTestCase):
    check = "demo cleanup regression"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_when_ownership_is_never_armed(self) -> None:
        self.repo.replace(BUILDER, "DidUnprotect = True", "DidUnprotect = False")
        self.assertFails("ownership must be armed only after Unprotect returns")


class PublicApiManifest(GateTestCase):
    check = "public API manifest"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_on_an_unrecorded_public_member(self) -> None:
        self.repo.replace(RUNTIME, "Option Private Module\r\n",
                          "Option Private Module\r\n\r\n"
                          "Public Function UI_TestOnlyProbe() As Boolean\r\nEnd Function\r\n")
        self.assertFails("[project-public] declared without a manifest entry",
                         "UI_TestOnlyProbe")


class PublicApiSelfTest(GateTestCase):
    check = "public API self-test"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_when_parameters_are_not_modelled(self) -> None:
        with patch.object(vba_api, "canonical_parameter", lambda text: "x"):
            self.assertFails("produced an identical manifest")


class SupportedApiDeclaration(GateTestCase):
    check = "supported API declaration"
    ROW = "| Supported API contract | Unchanged from the `[baseline v1.1.2]` facade |"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_when_an_unchanged_facade_is_declared_changed(self) -> None:
        self.repo.replace("CHANGELOG.md", self.ROW, "| Supported API contract | Minor: adds a member |")
        self.assertFails("is identical to [baseline")

    def test_fails_without_a_declaration_row(self) -> None:
        self.repo.replace("CHANGELOG.md", self.ROW, "")
        self.assertFails("CHANGELOG.md has no '| Supported API contract | ... |' row")


class WorkflowPinPolicy(GateTestCase):
    check = "workflow pin policy"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_on_a_moving_tag(self) -> None:
        text = self.repo.read_text(WORKFLOW)
        pinned = re.search(r"actions/checkout@[0-9a-f]{40} # v[0-9.]+", text)
        assert pinned is not None
        self.repo.replace(WORKFLOW, pinned[0], "actions/checkout@v4")
        self.assertFails("actions/checkout@v4 is not pinned to a commit SHA")

    def test_fails_without_a_tag_trigger(self) -> None:
        self.repo.replace(WORKFLOW, "    tags:\n      - 'v*'\n", "")
        self.assertFails("no push trigger on tags 'v*'")


class WorkflowPolicySelfTest(GateTestCase):
    check = "workflow policy self-test"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_when_any_comment_counts_as_audited(self) -> None:
        with patch.object(check_repo, "AUDIT_COMMENT_RE", re.compile(r".*")):
            self.assertFails("a version series is not an audited version")


class WikiBadgeSelfTest(GateTestCase):
    check = "wiki badge self-test"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_when_navigation_pages_need_a_badge(self) -> None:
        with patch.object(wiki_badges, "NAVIGATION_PAGES", ()):
            self.assertFails("a navigation page needs no badge")


class ReleaseState(GateTestCase):
    check = "release state"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_on_stale_release_markers(self) -> None:
        self.repo.append("README.md", "\nRelease Candidate\n\n[ ] sign off\n")
        self.repo.append("CHANGELOG.md", "\n_Nothing yet._\n")
        self.assertFails("README.md still contains 'Release Candidate'",
                         "README.md contains an unchecked release checklist",
                         "CHANGELOG.md contains an unfilled placeholder section")

    def test_fails_on_a_malformed_version(self) -> None:
        self.repo.write("VERSION", "v1.1.2\n")
        self.assertFails("VERSION does not hold a release number: 'v1.1.2'")

    def test_fails_when_version_is_not_the_newest_release(self) -> None:
        self.repo.write("VERSION", "1.1.3\n")
        self.assertFails("VERSION is 1.1.3, but the newest release in CHANGELOG.md is [1.1.2]")

    def test_fails_when_unreleased_is_not_the_first_version_section(self) -> None:
        self.repo.replace("CHANGELOG.md", "## [Unreleased]\n", "## [Unreleased] - 2026-11-01\n")
        self.assertFails("the first version section must be '## [Unreleased]'",
                         "malformed release heading '## [Unreleased] - 2026-11-01'")

    def test_fails_on_a_release_link_to_another_repository(self) -> None:
        self.repo.replace("CHANGELOG.md", "[1.0.1]: https://github.com/danielep71/VBA-EXCEL_UI/",
                          "[1.0.1]: https://github.com/example/VBA-EXCEL_UI/")
        self.assertFails("link reference [1.0.1] points outside this repository")


class ReleaseStateSelfTest(GateTestCase):
    check = "release state self-test"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_when_any_version_is_accepted(self) -> None:
        with patch.object(check_repo, "VERSION_RE", re.compile(r".*")):
            self.assertFails("a leading v is rejected")


class ReviewFindingCrosswalk(GateTestCase):
    check = "review finding crosswalk"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_on_an_unmapped_identifier(self) -> None:
        self.repo.append("docs/RIBBON_SDI_BEHAVIOR.md", "\nSee ICR-UI-112-P9-99.\n")
        self.assertFails("ICR-UI-112-P9-99 is cited in docs/RIBBON_SDI_BEHAVIOR.md but has no row")

    def test_fails_on_a_duplicate_row(self) -> None:
        row = next(line for line in self.repo.read_text("CHANGELOG.md").splitlines()
                   if line.startswith("| `ICR-UI-P1-01` |"))
        self.repo.replace("CHANGELOG.md", row + "\n", row + "\n" + row + "\n")
        self.assertFails("CHANGELOG.md crosswalk: ICR-UI-P1-01 has more than one row")

    def test_fails_when_a_renumbered_identifier_points_nowhere(self) -> None:
        self.repo.replace("CHANGELOG.md", "| `ICR-UI-112-P3-04` | v1.1.2", "| `ICR-UI-112-P3-44` | v1.1.2")
        self.assertFails("ICR-UI-111-P3-07 points at ICR-UI-112-P3-04, which has no row")

    def test_fails_without_the_crosswalk_section(self) -> None:
        self.repo.replace("CHANGELOG.md", "### 🔎 Review finding crosswalk", "### Review findings")
        self.assertFails("CHANGELOG.md has no '### 🔎 Review finding crosswalk' section")


class RepositoryHygiene(GateTestCase):
    check = "repository hygiene"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_on_a_tracked_workbook(self) -> None:
        self.repo.write("release.xlsm", b"PK\x03\x04")
        self.assertFails("release.xlsm: macro-enabled Excel workbook must not be tracked")

    def test_fails_when_a_blocked_pattern_leaves_gitignore(self) -> None:
        self.repo.replace(".gitignore", "*.xlsm\n", "")
        self.assertFails(".gitignore is missing required blocked pattern '*.xlsm'")

    def test_fails_on_a_tracked_file_the_gitignore_hides(self) -> None:
        self.repo.write("notes.tmp", "scratch\n")
        self.assertFails("notes.tmp: tracked file is hidden by the current .gitignore")

    # Binary content (#82). A ZIP local-file header carries NUL bytes, as does
    # every Office Open XML workbook, which is a ZIP container.
    ZIP = b"PK\x03\x04\x14\x00\x06\x00\x08\x00\x00\x00!\x00[Content_Types].xml"
    ALLOWLIST = "tools/binary_allowlist.txt"

    def test_fails_on_a_tracked_archive(self) -> None:
        self.repo.write("tools/bundle.zip", self.ZIP)
        self.assertFails("tools/bundle.zip: binary file outside images/ "
                         "(Git treats it as binary (text attribute unset))")

    def test_fails_on_a_workbook_renamed_with_a_text_extension(self) -> None:
        self.repo.write("docs/notes.txt", self.ZIP)
        self.assertFails("docs/notes.txt: binary file outside images/ (its content contains a NUL byte)")

    def test_fails_on_a_nul_byte_in_a_module(self) -> None:
        self.repo.replace(RUNTIME, "Option Explicit", "Option\0Explicit")
        self.assertFails(f"{RUNTIME}: binary file outside images/ (its content contains a NUL byte)")

    def test_fails_on_an_image_outside_images(self) -> None:
        self.repo.write("docs/diagram.png", self.repo.read_bytes("images/excel_ui_banner.png"))
        self.assertFails("docs/diagram.png: binary file outside images/")

    def test_passes_with_a_new_image_under_images(self) -> None:
        self.repo.write("images/diagram.png", self.repo.read_bytes("images/excel_ui_banner.png"))
        self.assertPasses()

    def test_passes_with_an_allowlisted_binary_while_its_digest_matches(self) -> None:
        self.repo.write("tools/bundle.zip", self.ZIP)
        self.repo.append(self.ALLOWLIST, f"{hashlib.sha256(self.ZIP).hexdigest()}  tools/bundle.zip\n")
        self.assertPasses()

    def test_fails_when_an_allowlisted_binary_changes(self) -> None:
        self.repo.write("tools/bundle.zip", self.ZIP + b"\0changed")
        self.repo.append(self.ALLOWLIST, f"{hashlib.sha256(self.ZIP).hexdigest()}  tools/bundle.zip\n")
        self.assertFails("tools/bundle.zip: allowlisted binary has changed")

    def test_fails_on_an_allowlist_entry_for_an_untracked_path(self) -> None:
        self.repo.append(self.ALLOWLIST, f"{'0' * 64}  tools/gone.zip\n")
        self.assertFails(f"{self.ALLOWLIST}: tools/gone.zip is listed but not tracked")

    def test_fails_on_a_malformed_allowlist_entry(self) -> None:
        self.repo.append(self.ALLOWLIST, "tools/bundle.zip\n")
        self.assertFails(f"{self.ALLOWLIST}:", "expected '<sha256>  <path>'")


class MarkdownLinks(GateTestCase):
    check = "markdown links"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_on_a_missing_target(self) -> None:
        self.repo.append("docs/RIBBON_SDI_BEHAVIOR.md", "\n[gone](missing-file.md)\n")
        self.assertFails("docs/RIBBON_SDI_BEHAVIOR.md: missing link target missing-file.md")

    def test_fails_on_a_broken_anchor(self) -> None:
        self.repo.append("docs/RIBBON_SDI_BEHAVIOR.md", "\n[gone](#no-such-heading)\n")
        self.assertFails("docs/RIBBON_SDI_BEHAVIOR.md: broken internal anchor #no-such-heading")


class MarkdownEscapes(GateTestCase):
    check = "markdown escapes"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_on_escaped_punctuation_in_prose(self) -> None:
        self.repo.append("docs/RIBBON_SDI_BEHAVIOR.md", "\nThe M\\_EXCEL\\_UI module.\n")
        self.assertFails("docs/RIBBON_SDI_BEHAVIOR.md:", "escaped Markdown outside code: \\_, \\_")


class MarkdownEscapeSelfTest(GateTestCase):
    check = "markdown escape self-test"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_when_no_escape_is_recognised(self) -> None:
        with patch.object(check_repo, "ESCAPE_RE", NEVER):
            self.assertFails("an escaped underscore is reported")


class LabelCatalogue(GateTestCase):
    check = "label catalogue"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_when_a_documented_colour_differs(self) -> None:
        self.repo.replace("docs/ISSUE_LABELS.md", "| `ribbon` | `#D4C5F9` |", "| `ribbon` | `#D4C5FA` |")
        self.assertFails("docs/ISSUE_LABELS.md catalogue table: missing or different: ribbon")

    def test_fails_when_the_documented_count_differs(self) -> None:
        self.repo.replace("docs/ISSUE_LABELS.md", "defines 27 issue labels", "defines 26 issue labels")
        self.assertFails("docs/ISSUE_LABELS.md does not state 27 labels")


class PrivateReviewDenylist(GateTestCase):
    check = "private review denylist"
    REVIEW = "docs/INDEPENDENT_CODE_REVIEW_V9.9.9.md"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_on_a_tracked_review(self) -> None:
        self.repo.write(self.REVIEW, "# Review\n")
        self.assertFails(f"{self.REVIEW} is tracked, but")

    def test_fails_on_a_review_present_in_the_working_tree(self) -> None:
        self.repo.write(self.REVIEW, "# Review\n", track=False)
        self.assertFails(f"{self.REVIEW} is present in the working tree")


class PrivateReviewDenylistSelfTest(GateTestCase):
    check = "private review denylist self-test"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_when_nothing_is_denied(self) -> None:
        with patch.object(check_repo, "PRIVATE_PATTERNS", []):
            self.assertFails("should have been denied")


class HouseStyleFormatter(GateTestCase):
    check = "house-style formatter"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_on_formatting_drift(self) -> None:
        self.repo.append(RUNTIME, "\r\n\r\n")
        self.assertFails(f"{RUNTIME}: not in house-style normal form")


class FormatterSelfTest(GateTestCase):
    check = "formatter self-test"

    def test_passes_on_the_current_tree(self) -> None:
        self.assertPasses()

    def test_fails_when_jumps_are_not_renamed(self) -> None:
        with patch.object(reformat, "rename_label_text", lambda text: text):
            self.assertFails("a real jump is still renamed")


if __name__ == "__main__":
    unittest.main()
