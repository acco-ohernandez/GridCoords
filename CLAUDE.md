Subject: C# Revit API — "GridCoords" Add-in

Project:  C:\Visual Studio Files\GridCoords
Entry:    GridCoords\Cmd_GridCoords.cs (IExternalCommand)
WPF:      GridCoords\Forms\GridCoords_Form.xaml (modeless)
Handler:  GridCoords\Common\GridCoords_EventHandler.cs (new)
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
    ├─ Constructor receives: UIApplication, ExternalEvent, handler
    ├─ Loaded event: position window center-right of Revit
    ├─ ComponentDispatcher.ThreadIdle: detect view changes,
    │   refresh grid list + update title/status with view name
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
  - RadioButton group: (H,V) or (V,H) order
  - Live preview TextBlock showing example with actual grid names

[Expander: Placement Options — collapsed by default]
  - Checkbox: "Auto offset" (default ON, scales with View.Scale)
  - Offset X / Offset Y numeric inputs (paper inches, converted
    to model units: modelOffset = paperInches * view.Scale / 12)
  - Manual inputs disabled when Auto is checked

[Expander: Existing Labels — collapsed by default]
  - RadioButton group:
    • Delete existing labels first (default)
    • Skip intersections that already have a label
    • Place duplicates anyway
  - Identification via Extensible Storage schema on each element

── RIGHT COLUMN: Grids + Results ──────────────────────────────

[Expander: Grids — always expanded, header includes view name]
  - Refresh button in expander header area
  - Auto-refreshes via ThreadIdle when active view changes
  
  DataGrid with columns:
    • Checkbox (toggle-all header)
    • Grid Name (natural sort via custom IComparer<string>)
    • Orientation (auto-detected: Horizontal | Vertical | Angled)
      — editable ComboBox so user can override classification
    
  Grid classification logic:
    • Use gridDir.DotProduct(view.RightDirection):
      abs ≥ 0.985 (cos 10°) → Horizontal
      abs ≤ 0.174 (sin 10°) → Vertical
      else → Angled (excluded by default, user can override)
    • Curved grids → labeled "Curved", excluded by default
    
  Below DataGrid:
    • TextBlock: "{N} intersections found" (live update)

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
  • Delete Labels → ExternalEvent: query ES schema in current 
    view, delete found elements, update Results
  • Close → close form

Status TextBlock below: "Ready." / "Working…" / "12 placed."

── CORE LOGIC (inside ExternalEvent HandlerAction) ────────────

1. Collect checked grids, separate into H and V groups
2. For every H×V pair: Curve.Intersect(otherCurve, out results)
   — check SetComparisonResult, extract XYZ from results
3. Filter to crop region: if view.CropBoxActive, test point 
   against view.CropBox BoundingBoxXYZ min/max (simple rect)
4. Handle existing labels per user setting (query ES schema)
5. Place TextNote.Create(doc, viewId, xyz+offset, label, typeId)
   or doc.Create.NewFamilyInstance(xyz, symbol, view) + set param
6. Tag each placed element with Extensible Storage:
     Schema GUID: (generate one stable GUID)
     Fields: ViewId (int), HGridName (string), VGridName (string)
7. Dispatcher.BeginInvoke → update Results panel + status

── TRANSACTION & ERROR HANDLING ───────────────────────────────

- Single Transaction wrapping all placements
- Validate before running:
    • Active view is plan/ceiling plan/area plan
    • At least one H and one V grid are checked
    • If FamilyInstance mode: symbol loaded + parameter exists
- All errors surface in Results panel — no modal TaskDialogs
  (form is modeless, stay non-blocking)

── FILES TO CREATE/MODIFY ─────────────────────────────────────

  Modify: Cmd_GridCoords.cs (static form, ExternalEvent setup)
  Modify: Forms\GridCoords_Form.xaml (full layout)
  Modify: Forms\GridCoords_Form.xaml.cs (code-behind + logic)
  Create: Common\GridCoords_EventHandler.cs
  Modify: Common\Utils.cs (add PositionWindowCenterRight)

── CODE QUALITY ───────────────────────────────────────────────

- Code-behind pattern (no MVVM — matches AutoTagger)
- Multi-version safe: use TextNote.Create 5-param overload,
  standard Extensible Storage API, all stable 2022–2026
- Natural sort comparer for grid names
- Clean ThreadIdle unsubscribe on form close
