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
  - ComboBox selections are preserved across Grid Scope changes
    and view refreshes (saved/restored by ElementId/name)

[Expander: Label Format]
  - TextBox with token template, default: ({H},{V})
    Tokens: {H} = "more horizontal" grid name,
            {V} = "more vertical" grid name
  - Empty template falls back to ({H},{V})
  - RadioButton group: (H,V) or (V,H) order
  - Live preview TextBlock using first enabled pairing's grids
  - Comprehensive tooltip on Template label with format examples
    and diagonal token assignment explanation

[Expander: Placement Options]
  Label Alignment section:
  - Vertical: Above grid (default) | Below grid
    Above = bottom edge of text touches horizontal grid
    Below = top edge of text touches horizontal grid
  - Horizontal: Right of grid (default) | Left of grid
    Right = left edge of text touches vertical grid
    Left = right edge of text touches vertical grid
  - Checkbox: "Rotate labels on diagonal grids" (default ON)
    Rotates text to align with the H-token grid direction

  Offset section:
  - Checkbox: "Auto" (default ON, scales padding with View.Scale)
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

Right column wrapped in ScrollViewer for overflow handling.

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

[Expander: Horizontal Grids — XAML-defined]
  - Header shows "(N of M)" selection count
  - Select All checkbox
  - WrapPanel of checkboxes (compact layout, natural sort)

[Expander: Vertical Grids — XAML-defined]
  - Same layout as Horizontal Grids

[Dynamic Group Expanders — programmatic]
  - Generated for diagonal and curved grid groups
  - Same visual structure as H/V expanders (Select All + WrapPanel)
  - Each group has self-contained SelectAll via GridGroup class
  - Only shown when non-orthogonal grids exist in view

[Expander: Intersection Pairings]
  - Shows all cross-group pairings (N×(N-1)/2)
  - Each pairing has enable/disable checkbox
  - Labels show "{H} × {V} = count" with token assignment
  - Token rule: group with smaller angle → {H} ("more horizontal")
  - Live count updates when grid selections change

[Intersection Count]
  - Total across all enabled pairings with per-pairing breakdown

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

  • Place Labels → ExternalEvent: iterate enabled pairings,
    compute intersections per pairing, place labels
  • Delete Labels → ExternalEvent: scope-aware deletion,
    collects all grid names from all displayed groups into
    a single HashSet, only deletes matching labels
    (deletes all in All Visible mode)
  • Close → close form
  • Buttons disable during execution, re-enable on complete

Status TextBlock below: "Ready." / "Placing labels…" / "12 placed."

── ANGLE-BASED GRID GROUPING ─────────────────────────────────

Grid classification uses angle relative to view.RightDirection:
  - Compute angle: acos(|dot(lineDir, viewRight)|) → 0°–180°
  - Cluster by circular distance with 10° tolerance
  - Circular distance: min(|a-b|, 180-|a-b|) for wraparound
  - 0° ± 10° → Horizontal (standard), 90° ± 10° → Vertical
  - All other angles → Diagonal (~N°)
  - Curved grids → separate Curved group (angle = -1)

Data model:
  - GridGroup: collection, label, angle, classification, SelectAll
  - IntersectionPairing: two groups + enable flag + token assignment
  - PairingSnapshot: lightweight UI-thread capture for ExternalEvent
    includes group angles + IsOrthogonalPairing flag

Hybrid UI approach:
  - H and V groups use XAML-defined expanders (backward compatible)
  - Non-standard groups use DynamicGroupsContainer (StackPanel)
  - All groups stored in _allGridGroups list for unified logic

── SMART LABEL PLACEMENT ──────────────────────────────────────

BBox-based alignment (post-placement adjustment):
  1. Create element at exact intersection point
  2. doc.Regenerate() to ensure BoundingBox is computed
  3. Measure element bbox edge distances from intersection
  4. Compute alignment move vector:
     - Auto: align specified edge to grid line + 1/32" padding
     - Manual: apply user-specified X/Y offsets (axis-aligned)
  5. For diagonal grids (when rotation enabled):
     - Compute rotation angle from H-grid tangent via atan2
     - Constrain to ±90° for readability (never upside down)
     - Rotate alignment vector by same angle
     - Rotate element via ElementTransformUtils.RotateElement
  6. Move element via ElementTransformUtils.MoveElement

Works identically for TextNote and FamilyInstance.

── CORE LOGIC (inside ExternalEvent HandlerAction) ────────────

1. Snapshot enabled pairings with selected grid items from UI thread
2. For each enabled pairing, iterate H-token-group × V-token-group
3. Curve.Intersect(otherCurve, out results)
   — Handle multiple intersection points (curved grids)
   — check SetComparisonResult, extract XYZ from results
   — Curve.Intersect deprecated in 2026, suppressed via #pragma
4. Filter to crop region: if view.CropBoxActive, test point
   against view.CropBox BoundingBoxXYZ min/max (simple rect)
5. Handle existing labels per user setting (query ES schema)
6. Place element at intersection, measure bbox, align, rotate, move
7. Tag each placed element with Extensible Storage:
     Schema GUID: A1B2C3D4-E5F6-7890-ABCD-EF1234567890
     Fields: ViewIdStr (string), HGridName (string), VGridName (string)
     ViewId stored as string for multi-version compatibility
8. Dispatcher.BeginInvoke → update Results panel + status

── TRANSACTION & ERROR HANDLING ───────────────────────────────

- Single Transaction wrapping all placements
- Validate before running:
    • Active view is plan/ceiling plan/area plan
    • At least one enabled pairing with selected grids
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
  • DistributeGrids() → BuildDynamicGroupExpanders(), BuildPairingsUI()

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
- Fully-qualified WPF types in programmatic UI building:
    System.Windows.Controls.CheckBox (avoids WinForms ambiguity)
    System.Windows.Controls.Orientation
- Natural sort comparer for grid names
- Clean ThreadIdle unsubscribe on form close
- GridSelectionFilter (ISelectionFilter) for Pick Grids mode
- Very descriptive variable names throughout (no abbreviations)
