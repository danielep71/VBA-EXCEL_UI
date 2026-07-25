---
name: ✨ Feature request
about: Suggest a backward-compatible UI-control, recovery, diagnostic, demo, or testing enhancement
title: "[Feature]: "
labels: enhancement
---

## ✨ Problem / use case

Describe the real workflow or limitation.

What are you trying to achieve that the current `UI_...` API does not support
cleanly?

## 💡 Proposed behavior

Describe the desired outcome before prescribing implementation details.

Where relevant, provide a proposed call:

```vba
'Illustrative only
```

Explain:

```text
Requested behavior:
Expected scope:
Expected default:
Expected result or diagnostics:
```

## 🎯 Affected area

Check all that apply:

- [ ] Public `UI_...` API
- [ ] Tri-state semantics
- [ ] Ribbon control
- [ ] Application-level properties
- [ ] Window-level properties
- [ ] Title-bar / WinAPI handling
- [ ] Snapshot capture
- [ ] Snapshot restoration
- [ ] Structured diagnostics
- [ ] Emergency recovery
- [ ] Active-window or workbook targeting
- [ ] Demo workbook
- [ ] Regression harness
- [ ] Documentation
- [ ] Repository maintenance
- [ ] Other

## 🧭 Targeting and ownership

Describe the intended target:

- [ ] Current Excel application
- [ ] All open Excel windows
- [ ] Active window only
- [ ] Specified `Window`
- [ ] Active workbook only
- [ ] Specified `Workbook`
- [ ] Excel main window represented by `Application.Hwnd`
- [ ] Other

Explain how the feature should interact with UI changes made by other workbooks,
macros, or add-ins.

## 🔄 State and recovery behavior

Explain:

- whether prior state must be captured;
- whether the operation should be reversible;
- what should happen after a VBA project reset;
- what emergency recovery path should remain available;
- what should happen when a requested read or write fails.

## 📋 Diagnostic contract

Should the feature:

- [ ] remain fire-and-forget;
- [ ] return a Boolean result;
- [ ] add structured failure details;
- [ ] introduce machine-readable failure categories;
- [ ] log to the Immediate Window;
- [ ] preserve the current contract without new diagnostics?

Describe the desired contract.

## 🔀 Alternatives considered

Describe current workarounds or alternative designs, for example:

- composing existing `UI_SetExcelUI` calls;
- using `UI_CaptureExcelUIState` and `UI_ResetExcelUIToSnapshot`;
- direct Excel object-model property writes;
- a workbook-specific macro;
- a custom Ribbon or add-in;
- leaving the feature outside this repository.

## 🧩 Compatibility and Semantic Versioning

```text
Existing calls affected:
Backward compatible: Yes / No / Unsure
Suggested release:   patch / minor / major / unsure
```

Explain any proposed new public member, enum value, optional parameter, or changed
default.

## 🪟 Platform considerations

Identify any known concerns involving:

- Office 32-bit versus 64-bit;
- Excel versions;
- Windows versions;
- `Application.Hwnd`;
- Ribbon behavior;
- full-screen or maximized windows;
- other add-ins;
- Protected View;
- macro-security settings.

## 🧪 Validation proposal

How should the feature be verified?

Check all that apply:

- [ ] Core regression case
- [ ] Title-bar regression case
- [ ] Multi-window test
- [ ] Failure-path test
- [ ] `ScreenUpdating` preservation
- [ ] Manual recovery test
- [ ] Demo update
- [ ] 32-bit Office validation
- [ ] 64-bit Office validation
- [ ] Multiple Excel-version validation

Add proposed test cases or expected behavior.

## 📎 Additional context

Include sanitized screenshots, pseudocode, links, or examples that clarify the
request.
