# FAS Move Mode Grid and HUD Design

**Date:** 2026-07-31  
**Status:** Ready for user review  
**Repositories:** `Instrumenta` and `Instrumenta-Keys`

## Summary

FAS Move Mode will become a keyboard-first geometry tool built around three
explicit coordinate levels:

1. An existing named PowerPoint zone, normally the Content zone.
2. A configurable grid of rows and columns inside that zone.
3. One active grid field, with a 3 x 3 numpad anchor system inside it.

The current large VBA palette will no longer be the primary interface. A small,
non-activating Windows HUD will show the grid and active field, while a
click-through overlay will outline the active field on the slide. Neither HUD
element may modify the presentation or enter PowerPoint's undo history.

## Goals

- Let a user move a selected object to an anchor inside any grid field without
  using the mouse.
- Let a user width-fit one selected object to the active field while preserving
  aspect ratio and then place it at a chosen anchor.
- Make rows, columns, active field, and anchor visibly understandable.
- Preserve the current arrow-key modifier logic for fine/coarse movement and
  resizing.
- Keep the HUD compact and away from the centre of the slide.
- Keep all visual feedback outside the PowerPoint slide object model.

## Non-goals

- Timed fading of the field overlay. Version one keeps the overlay visible while
  Move Mode is active; fading can be added later without changing geometry.
- Assigning `Shift + Numpad 1-9`, `Ctrl + Shift + Numpad 1-9`, Alt combinations,
  Numpad Decimal, or Numpad Enter. These combinations remain unassigned until a
  concrete workflow requires them.
- Replacing ordinary PowerPoint alignment and distribution commands.
- Fitting multiple selected shapes. Moving treats a multi-shape selection as one
  bounding box; width-fit continues to require exactly one shape.
- A mouse-driven configuration panel in version one.

## Terminology

- **Zone:** An existing Instrumenta layout rectangle such as Content, Title, or
  Subtitle.
- **Grid:** The active zone divided into 1-4 rows and 1-4 columns.
- **Field:** One grid cell. Fields are numbered in row-major order from the
  top-left.
- **Active field:** The field used by movement, fitting, the HUD, and the overlay.
- **Anchor:** One of nine positions inside the active field, matching the spatial
  arrangement of Numpad 1-9.

For a 2 x 2 grid, fields are numbered as follows:

```text
+---------+---------+
| field 1 | field 2 |
+---------+---------+
| field 3 | field 4 |
+---------+---------+
```

If field 4 is active, Numpad 7 means the top-left anchor inside field 4, not the
top-left of the entire Content zone.

## Grid Geometry

The grid is calculated inside the active zone using a fixed 10 pt gutter between
adjacent fields in both directions.

```text
fieldWidth  = (zoneWidth  - ((columns - 1) * 10 pt)) / columns
fieldHeight = (zoneHeight - ((rows - 1) * 10 pt)) / rows
```

Valid row and column counts are 1, 2, 3, and 4. A 1 x 1 grid represents the full
active zone and is the default for a new installation. Saved valid row and
column preferences are restored when Move Mode starts. The active field resets
to row 1, column 1 on each activation and is then retained across slides until
Move Mode exits. Grid calculations and placement must reuse Instrumenta's
rotation-aware `GetRealLeft`, `GetRealTop`, `GetRealWidth`, `GetRealHeight`,
`SetRealLeft`, `SetRealTop`, and `SetRealWidth` helpers.

Changing the row or column count clamps the active row and column to the nearest
valid field. The active field is retained across slides while Move Mode remains
active.

## Keyboard Contract

The numpad is interpreted only when PowerPoint is foreground, the active
PowerPoint process matches the Move Mode session PID, and Move Mode is active.

| Input | Action |
|---|---|
| `Numpad 0` | Toggle Move Mode |
| `Numpad /` | Cycle rows: 1 -> 2 -> 3 -> 4 -> 1 |
| `Numpad *` | Cycle columns: 1 -> 2 -> 3 -> 4 -> 1 |
| `Numpad -` | Select previous field, wrapping in row-major order |
| `Numpad +` | Select next field, wrapping in row-major order |
| `Numpad 1-9` | Move selection to that anchor inside the active field, preserving size |
| `Ctrl + Numpad 1-9` | Width-fit one shape to the active field, preserve aspect ratio, then place at that anchor |
| Arrow keys | Fine movement using the existing 1 pt preference |
| `Shift + Arrow` | Coarse movement using the existing 10 pt preference |
| `Ctrl + Arrow` | Fine width/height resizing |
| `Ctrl + Shift + Arrow` | Coarse width/height resizing |
| `Shift + Numpad +` | Bring forward |
| `Shift + Numpad -` | Send backward |
| `Ctrl + Shift + Numpad +` | Bring to front |
| `Ctrl + Shift + Numpad -` | Send to back |
| `Esc` | Exit Move Mode |

Unassigned modifier combinations are not intercepted and pass through to
PowerPoint. Alt-based numpad combinations are deliberately avoided because
Windows uses Alt+numpad for character entry.

## Movement Semantics

Plain Numpad 1-9 moves the selection without resizing it. The selection's real
bounding box is aligned to the requested anchor inside the active field:

```text
7 = top-left      8 = top-centre      9 = top-right
4 = middle-left   5 = centre          6 = middle-right
1 = bottom-left   2 = bottom-centre   3 = bottom-right
```

Multiple selected shapes move together while preserving their internal spacing.
If the selection is larger than the active zone, the command does nothing and
the HUD reports the reason. A selection may be larger than the active field; in
that case the requested anchor is still honoured and the selection may extend
outside that field, because the command is movement rather than fitting.

## Width-fit Semantics

`Ctrl + Numpad 1-9` is the preset resizing layer. It operates on exactly one
selected shape:

1. Set the shape's real width to the active field width.
2. Preserve the shape's aspect ratio when `LockAspectRatio = msoTrue`.
3. Do not shrink the shape merely because its derived height exceeds the active
   field height. Width is the authoritative sizing rule.
4. If the derived height exceeds the entire active zone height, reduce the shape
   proportionally until it fits the active zone height.
5. Place the fitted shape at the requested numpad anchor inside the active field.

For unlocked shapes, version one changes width only and retains existing height.
It does not distort the shape to fill both field dimensions.

## State Model and PowerPoint Interface

`ModuleFASEditMode.bas` remains the authoritative owner of layout state. It will
publish the following values under the existing
`HKCU\\Software\\VB and VBA Program Settings\\InstrumentaFASEditMode` registry
area so the companion can render without scraping the VBA form:

### Preferences

- `GridRows` (1-4)
- `GridColumns` (1-4)

### Session

- `Active`
- `PowerPointPid`
- `SessionId`
- `ActiveRow`
- `ActiveColumn`
- `ZoneName`
- `ZoneLeftPt`
- `ZoneTopPt`
- `ZoneWidthPt`
- `ZoneHeightPt`
- `FieldLeftPt`
- `FieldTopPt`
- `FieldWidthPt`
- `FieldHeightPt`
- `StatusText`
- `StateRevision`, incremented whenever visible state changes

The VBA module exposes focused public macros for cycling rows, cycling columns,
selecting the previous/next field, moving to each anchor, and width-fitting to
each anchor. Every state-changing macro republishes the session snapshot after
completing or rejecting the action. Snapshot fields are written first and
`StateRevision` is written last. The companion reads the revision before and
after the remaining values and accepts the snapshot only when both revision
reads match, preventing a partially updated HUD.

## Compact HUD

The primary interface is a small borderless Windows Forms HUD hosted by
`Instrumenta-Keys`.

- Approximate size: 150 x 80 logical pixels.
- Position: top-right of the PowerPoint editing workspace with an 18 px inset.
- Behaviour: topmost relative to PowerPoint, non-activating, and click-through.
- Visibility: shown only when Move Mode is active and the matching PowerPoint
  process is foreground.
- Version one uses a fixed calculated position; dragging and saved custom
  positions are deferred.

The HUD displays:

```text
MOVE MODE
GRID 2 x 2   FIELD 4/4
[mini-grid with field 4 highlighted]
Top-left
```

The mini-grid uses an orange fill for the active field. The final line shows the
most recent anchor or status message, including errors such as `Select one
shape to fit`.

The existing large `FASEditModeForm` is not shown automatically when Move Mode
starts. It may remain temporarily as a development fallback until the HUD passes
the live smoke test, after which normal activation uses only the companion HUD.

## Active-field Overlay

`Instrumenta-Keys` hosts a second borderless, transparent, click-through Windows
Forms overlay positioned over the active field:

- 2 px orange border.
- Approximately 6% orange fill opacity.
- No caption, taskbar entry, activation, or mouse interception.
- Persistent while Move Mode is active and matching PowerPoint is foreground.
- Hidden immediately on mode exit, PID mismatch, PowerPoint losing foreground,
  unavailable slide geometry, slide show view, or companion shutdown.

The overlay must never be represented by a PowerPoint shape. It therefore cannot
alter the slide, selection, saved file, or undo stack.

The companion reads the field rectangle in slide points from the registry and
uses the active PowerPoint `DocumentWindow.PointsToScreenPixelsX` and
`PointsToScreenPixelsY` methods to calculate screen pixels. It refreshes the
screen rectangle at a moderate interval only while Move Mode is active, so zoom,
window movement, task panes, and slide changes reposition the overlay. A failed
coordinate refresh hides the overlay rather than displaying it in a stale
location.

## Component Boundaries

### Instrumenta PowerPoint add-in

- Owns zone, grid, field, anchor, movement, and fitting calculations.
- Publishes an atomic state snapshot through the existing registry contract.
- Does not create HUD or overlay shapes.
- Continues to use PowerPoint events to refresh selection and zone state.

### Instrumenta Keys companion

- Owns low-level key routing.
- Reads the published Move Mode snapshot.
- Owns the compact HUD and active-field overlay windows.
- Converts slide points to screen pixels through the active PowerPoint window.
- Hides all Move Mode UI whenever safety checks fail.

The companion does not duplicate grid geometry formulas. It renders the field
rectangle calculated and published by the add-in.

## Failure Handling

- No selected shapes: movement/fitting is rejected and `StatusText` explains the
  required selection.
- More than one shape for width-fit: fitting is rejected; movement remains
  available.
- Invalid saved grid values: clamp to 1-4 and republish corrected state.
- Active field outside a changed grid: clamp row/column before calculating the
  field rectangle.
- No active presentation, unsupported view, COM failure, or stale PID: hide HUD
  and overlay; never display geometry from another PowerPoint process.
- Companion not running: VBA geometry commands still work; only the HUD and
  overlay are absent.
- VBA add-in not available: the companion passes unhandled keys through and does
  not show Move Mode UI.

## Testing Strategy

### Pure key-routing tests

Extend `Instrumenta-Keys/tests/test-edit-mode-routing.ps1` to cover:

- `/`, `*`, `-`, and `+` with no modifiers.
- All nine plain anchor mappings.
- All nine Ctrl width-fit mappings.
- Existing arrow, z-order, PID, foreground, NumLock-off, and Alt pass-through
  behaviour.
- Unassigned Shift+numpad digit pass-through behaviour.

### VBA source-contract tests

Extend `Instrumenta/scripts/test-fas-edit-mode.ps1` to verify:

- Grid range accepts 1-4.
- Active row/column state and cycle macros exist.
- Field calculations include 10 pt horizontal and vertical gutters.
- Width-fit does not use field height as its normal limiting dimension.
- State snapshot keys are written.
- Existing rotation-aware real-bound helpers remain in use.

### Companion HUD tests

Add pure helper tests for:

- Row-major field indexing and labels.
- Mini-grid rendering state.
- HUD/overlay visibility decisions for Active, foreground, PID match, and valid
  geometry combinations.
- Screen rectangle conversion failures hiding the overlay.

### Live smoke test

Automate a temporary PowerPoint instance and verify:

1. Numpad 0 activates Move Mode.
2. A 2 x 2 grid can select field 4.
3. Numpad 7 moves a shape to the top-left anchor inside field 4.
4. Ctrl+Numpad 7 width-fits one locked-aspect shape to field 4 and preserves its
   expected aspect ratio.
5. The registry snapshot reports grid 2 x 2, field 4, and the correct field point
   bounds.
6. The companion reports HUD and overlay visible while PowerPoint is foreground.
7. Exiting Move Mode clears the active session and hides both windows.

Visual verification at 50%, 100%, and 150% PowerPoint zoom confirms that the
overlay remains aligned with the active field. Production installation follows
only after source-contract tests, routing tests, build, COM smoke, process-count
check, backup, and installed-file hash verification pass.

## Delivery Order

1. Implement and validate grid/field state, keyboard routing, movement, and
   width-fit semantics without changing the current palette lifecycle.
2. Add the compact companion HUD and switch normal activation away from the
   large VBA form after the HUD passes live smoke.
3. Add and validate the persistent screen-aligned field overlay.
4. Build, back up, install, and smoke-test both the PPAM and Instrumenta Keys
   companion.

Timed fading, draggable HUD positioning, and additional modifier layers remain
separate follow-up enhancements.
