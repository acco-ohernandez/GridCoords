Subject: C# Revit API — "GridCoords" Add-in

Project:  C:\Visual Studio Files\GridCoords
Entry:    GridCoords\Cmd_GridCoords.cs (IExternalCommand)
WPF:      GridCoords\Forms\GridCoords_Form.xaml (modeless)
Handler:  GridCoords\Common\GridCoords_EventHandler.cs
Target:   Revit 2022–2026 (multi-target via csproj)

── ARCHITECTURE ───────────────────────────────────────────────

Follow the exact modeless pattern from the AutoTagger:

  Cmd_GridCoords (IExternalCommand)
    ├─ static GridCoords_Form _form  (single-instance guard)
    ├─ Execute(): create handler + ExternalEvent, pass to form
    ├─ If form already visible → reactivate, return
    └─ Collect initial grid data on Revit thread, pass to form

  GridCoords_EventHandler (IExternalEventHandler)
    ├─ Action<UIApplication> HandlerAction { get; set; }
    ├─ Execute(): invoke action, catch OperationCanceledException
    └─ Finally: clear HandlerAction = null

  GridCoords_Form (modeless WPF Window)
    ├─ Constructor receives: UIApplication, ExternalEvent, handler,
    │   initial grid data, view name, view ID
    ├─ Loaded event: position window center-right of Revit
    ├─ ComponentDispatcher.ThreadIdle: detect view changes +
    │   selection changes (for "Selected Grids" scope mode)
    ├─ All Revit API writes go through ExternalEvent.Raise()
    └─ Closed event: unsubscribe ThreadIdle, dispose

── MODELESS FORM LAYOUT ───────────────────────────────────────

Topmost=True, ResizeMode=CanResize, WindowStartupLocation=Manual
Two-column layout with Expander sections.
Style resources matching AutoTagger (Expander, RadioButton,
CheckBox, Button, NumericInput styles).

── LEFT COLUMN: Settings ──────────────────────────────────────

[Expander: Placement Mode]
  - RadioButton group: TextNote | FamilyInstance
  - If TextNote → ComboBox: available TextNoteTypes in doc
  - If FamilyInstance → ComboBox: loaded Generic Annotation
    families/types + ComboBox: text parameters on selected type
  - FamilySymbol activation handled before placement:
    if (!symbol.IsActive) symbol.Activate(); doc.Regenerate();

[Expander: Label Format]
  - TextBox with token template, default: ({H},{V})
    Tokens: {H} = horizontal grid name, {V} = vertical grid name
  - Empty template falls back to ({H},{V})
  - RadioButton group: (H,V) or (V,H) order
  - Live preview TextBlock showing example with actual grid names
  - Comprehensive tooltip on Template label with format examples

[Expander: Placement Options]
  - Checkbox: "Auto offset" (default ON, scales with View.Scale)
  - Offset X / Offset Y numeric inputs (paper inches, converted
    to model units: modelOffset = paperInches * view.Scale / 12)
  - Manual inputs disabled when Auto is checked

[Expander: Existing Labels]
  - RadioButton group:
    • Delete existing labels first (default)
    • Skip intersections that already have a label
    • Place duplicates anyway
  - Identification via Extensible Storage schema on each element

── RIGHT COLUMN: Grids + Results ──────────────────────────────

[View Info Bar]
  - View name (italic, gray) + Refresh button
  - Auto-refreshes via ThreadIdle when active view changes

[Expander: Grid Scope]
  - RadioButton group:
    • All Visible Grids (default) — every grid in the view
    • Currently Selected Grids — filters to Revit selection,
      auto-refreshes via ThreadIdle on selection change
    • Pick Grids — Pick/Clear buttons, minimizes form for
      PickObjects with GridSelectionFilter (ISelectionFilter),
      shows "{N} picked" count
  - Scope affects both Place Labels and Delete Labels

[Expander: Horizontal Grids]
  - Header shows "(N of M)" selection count
  - Select All checkbox
  - WrapPanel of checkboxes (compact layout, natural sort)

[Expander: Vertical Grids]
  - Same layout as Horizontal Grids

[Expander: Angled / Curved — collapsed, hidden when empty]
  - Only shown when angled/curved grids exist in view
  - Each grid shows detected orientation + ComboBox to
    reassign as Horizontal or Vertical

[Intersection Count]
  - "{N} intersections  (H x V)" with live math breakdown

[Expander: Results — starts Collapsed, shown after execution]
  Grid layout:
    • Labels placed: {N}
    • Labels deleted: {N}
    • Skipped: {N}
    • Errors: {N} (red if > 0)
    • Detail TextBlock (italic, gray, collapsed unless errors)

── BOTTOM ACTION BAR ──────────────────────────────────────────

Three buttons centered:
  [Place Labels]  [Delete Labels]  [Close]

  • Place Labels → ExternalEvent: run intersection logic, place
  • Delete Labels → ExternalEvent: scope-aware deletion,
    only deletes labels matching displayed grids when using
    Selected/Pick scope (deletes all in All Visible mode)
  • Close → close form
  • Buttons disable during execution, re-enable on complete

Status TextBlock below: "Ready." / "Placing labels…" / "12 placed."

── CORE LOGIC (inside ExternalEvent HandlerAction) ────────────

1. Collect checked grids from H and V lists (+ reassigned Others)
2. For every H×V pair: Curve.Intersect(otherCurve, out results)
   — check SetComparisonResult, extract XYZ from results
   — Curve.Intersect deprecated in 2026, suppressed via #pragma
3. Filter to crop region: if view.CropBoxActive, test point
   against view.CropBox BoundingBoxXYZ min/max (simple rect)
4. Handle existing labels per user setting (query ES schema)
5. Place TextNote.Create(doc, viewId, xyz+offset, label, typeId)
   or doc.Create.NewFamilyInstance(xyz, symbol, view) + set param
6. Tag each placed element with Extensible Storage:
     Schema GUID: A1B2C3D4-E5F6-7890-ABCD-EF1234567890
     Fields: ViewIdStr (string), HGridName (string), VGridName (string)
     ViewId stored as string for multi-version compatibility
7. Dispatcher.BeginInvoke → update Results panel + status

── TRANSACTION & ERROR HANDLING ───────────────────────────────

- Single Transaction wrapping all placements
- Validate before running:
    • Active view is plan/ceiling plan/area plan
    • At least one H and one V grid are checked
    • If FamilyInstance mode: symbol loaded + parameter exists
    • Template field: falls back to ({H},{V}) if empty
    • TagElement: null-guarded against failed element creation
- All errors surface in Results panel — no modal TaskDialogs
  (form is modeless, stay non-blocking)

── WPF INITIALIZATION SAFETY ─────────────────────────────────

All methods called during constructor that touch XAML controls
must include null guards (controls don't exist until after
InitializeComponent completes XAML parsing):
  • UpdateFormatPreview()
  • UpdateExpanderHeaders()
  • UpdateIntersectionCount()
  • DistributeGrids()

── FILES ──────────────────────────────────────────────────────

  GridCoords\Cmd_GridCoords.cs          — IExternalCommand, static form
  GridCoords\Forms\GridCoords_Form.xaml — Two-column WPF layout
  GridCoords\Forms\GridCoords_Form.xaml.cs — Code-behind + all logic
  GridCoords\Common\GridCoords_EventHandler.cs — IExternalEventHandler
  GridCoords\Common\Utils.cs            — PositionWindowCenterRight,
                                          NaturalSortCompare

── CODE QUALITY ───────────────────────────────────────────────

- Code-behind pattern (no MVVM — matches AutoTagger)
- Multi-version safe: TextNote.Create 5-param overload,
  ElementId.ToString() instead of IntegerValue,
  standard Extensible Storage API, all stable 2022–2026
- Type aliases to avoid ambiguity:
    RevitGrid = Autodesk.Revit.DB.Grid
    WinVisibility = System.Windows.Visibility
- Natural sort comparer for grid names
- Clean ThreadIdle unsubscribe on form close
- GridSelectionFilter (ISelectionFilter) for Pick Grids mode
