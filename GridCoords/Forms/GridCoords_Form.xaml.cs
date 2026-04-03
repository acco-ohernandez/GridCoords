using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Threading;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI.Selection;
using GridCoords.Common;
using RevitGrid = Autodesk.Revit.DB.Grid;
using WinVisibility = System.Windows.Visibility;

namespace GridCoords.Forms
{
    public partial class GridCoords_Form : Window
    {
        // ── Revit references ──
        private readonly UIApplication _uiApp;
        private readonly ExternalEvent _externalEvent;
        private readonly GridCoords_EventHandler _handler;
        private IntPtr _revitHandle;

        // ── State: three separate lists for H, V, and Other grids ──
        private readonly ObservableCollection<GridRowItem> _hGridItems = new ObservableCollection<GridRowItem>();
        private readonly ObservableCollection<GridRowItem> _vGridItems = new ObservableCollection<GridRowItem>();
        private readonly ObservableCollection<GridRowItem> _otherGridItems = new ObservableCollection<GridRowItem>();
        private ElementId _lastViewId;
        private bool _suppressSelectAllEvents;

        // ── Scope state ──
        private string _pendingParameterRestore;
        private List<ElementId> _pickedGridIds = new List<ElementId>();
        private HashSet<ElementId> _lastSelectionIds = new HashSet<ElementId>();

        // ── Extensible Storage schema ──
        private static readonly Guid SchemaGuid = new Guid("A1B2C3D4-E5F6-7890-ABCD-EF1234567890");
        private const string SchemaName = "GridCoordsData";
        private const string FieldViewId = "ViewIdStr";
        private const string FieldHGridName = "HGridName";
        private const string FieldVGridName = "VGridName";

        public GridCoords_Form(UIApplication uiApp, ExternalEvent externalEvent,
                               GridCoords_EventHandler handler, List<GridRowItem> initialGrids,
                               string viewName, ElementId viewId)
        {
            InitializeComponent();

            _uiApp = uiApp;
            _externalEvent = externalEvent;
            _handler = handler;
            _revitHandle = uiApp.MainWindowHandle;
            _lastViewId = viewId;

            // Bind the three ItemsControls
            HGridList.ItemsSource = _hGridItems;
            VGridList.ItemsSource = _vGridItems;
            OtherGridList.ItemsSource = _otherGridItems;

            // Distribute grids into the three groups
            DistributeGrids(initialGrids);

            TblViewName.Text = $"View: {viewName}";
            UpdateExpanderHeaders();
            UpdateIntersectionCount();
            UpdateFormatPreview();

            // Populate dropdowns on Revit thread (we're still in Execute context)
            PopulateDropdowns(uiApp.ActiveUIDocument.Document);

            this.Loaded += OnLoaded;
            this.Closed += OnClosed;
        }

        // ── Grid distribution into H / V / Other ──

        private void DistributeGrids(List<GridRowItem> allGrids)
        {
            _hGridItems.Clear();
            _vGridItems.Clear();
            _otherGridItems.Clear();

            foreach (var g in allGrids)
            {
                switch (g.Orientation)
                {
                    case "Horizontal":
                        _hGridItems.Add(g);
                        break;
                    case "Vertical":
                        _vGridItems.Add(g);
                        break;
                    default: // Angled, Curved
                        _otherGridItems.Add(g);
                        break;
                }
            }

            // Show/hide Angled/Curved section
            if (ExpanderOtherGrids != null)
                ExpanderOtherGrids.Visibility = _otherGridItems.Count > 0
                    ? WinVisibility.Visible
                    : WinVisibility.Collapsed;
        }

        private void UpdateExpanderHeaders()
        {
            if (ExpanderHGrids == null || ExpanderVGrids == null || ExpanderOtherGrids == null) return;

            int hSelected = _hGridItems.Count(g => g.IsSelected);
            int vSelected = _vGridItems.Count(g => g.IsSelected);
            ExpanderHGrids.Header = $"Horizontal Grids ({hSelected} of {_hGridItems.Count})";
            ExpanderVGrids.Header = $"Vertical Grids ({vSelected} of {_vGridItems.Count})";

            if (_otherGridItems.Count > 0)
                ExpanderOtherGrids.Header = $"Angled / Curved ({_otherGridItems.Count})";
        }

        /// <summary>
        /// All grid items across all three groups (convenience accessor for placement logic).
        /// </summary>
        private IEnumerable<GridRowItem> AllGridItems =>
            _hGridItems.Concat(_vGridItems).Concat(_otherGridItems);

        // ── Lifecycle ──

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            Utils.PositionWindowCenterRight(this, _revitHandle);
            ComponentDispatcher.ThreadIdle += OnThreadIdle;
        }

        private void OnClosed(object sender, EventArgs e)
        {
            ComponentDispatcher.ThreadIdle -= OnThreadIdle;
        }

        private void OnThreadIdle(object sender, EventArgs e)
        {
            try
            {
                var uidoc = _uiApp.ActiveUIDocument;
                if (uidoc == null) return;

                // Detect view change → always refresh
                var currentViewId = uidoc.ActiveView?.Id;
                if (currentViewId != null && !currentViewId.Equals(_lastViewId))
                {
                    _lastViewId = currentViewId;
                    RefreshGridList();
                    return;
                }

                // In "Selected Grids" mode, detect selection changes
                if (RbScopeSelected?.IsChecked == true)
                {
                    var currentSelIds = new HashSet<ElementId>(uidoc.Selection.GetElementIds());
                    if (!currentSelIds.SetEquals(_lastSelectionIds))
                    {
                        _lastSelectionIds = currentSelIds;
                        RefreshGridList();
                    }
                }
            }
            catch
            {
                // Ignore — Revit may not be ready
            }
        }

        // ── Dropdown population ──

        private void PopulateDropdowns(Document doc)
        {
            // Save current selections by ElementId so we can restore them
            var prevTextTypeId = (CmbTextNoteTypes.SelectedItem as TextNoteType)?.Id;
            var prevFamilyTypeId = (CmbFamilyTypes.SelectedItem as FamilySymbol)?.Id;
            _pendingParameterRestore = (CmbParameters.SelectedItem as ParameterItem)?.Name;

            // TextNoteTypes
            var textTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(TextNoteType))
                .Cast<TextNoteType>()
                .OrderBy(t => t.Name)
                .ToList();
            CmbTextNoteTypes.ItemsSource = textTypes;
            // Restore previous selection, or default to first
            int restoredTextIdx = prevTextTypeId != null
                ? textTypes.FindIndex(t => t.Id.Equals(prevTextTypeId))
                : -1;
            CmbTextNoteTypes.SelectedIndex = restoredTextIdx >= 0 ? restoredTextIdx : (textTypes.Count > 0 ? 0 : -1);

            // Generic Annotation family types
            var annotationTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericAnnotation)
                .Cast<FamilySymbol>()
                .OrderBy(s => s.FamilyName)
                .ThenBy(s => s.Name)
                .ToList();
            CmbFamilyTypes.ItemsSource = annotationTypes;
            // Restore previous selection, or default to first
            int restoredFamilyIdx = prevFamilyTypeId != null
                ? annotationTypes.FindIndex(s => s.Id.Equals(prevFamilyTypeId))
                : -1;
            CmbFamilyTypes.SelectedIndex = restoredFamilyIdx >= 0 ? restoredFamilyIdx : (annotationTypes.Count > 0 ? 0 : -1);
        }

        private void PopulateParametersForFamily(Document doc, FamilySymbol symbol)
        {
            if (symbol == null) { CmbParameters.ItemsSource = null; return; }

            var paramItems = new List<ParameterItem>();

            var family = symbol.Family;
            if (family != null)
            {
                // Collect from symbol (type) parameters
                foreach (Parameter p in symbol.Parameters)
                {
                    if (p.StorageType == StorageType.String && !p.IsReadOnly)
                        paramItems.Add(new ParameterItem { Name = p.Definition.Name, ParamDefinition = p.Definition });
                }
                // Also get instance parameters from a placed instance (if any exist)
                var sampleInstance = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .OfCategory(BuiltInCategory.OST_GenericAnnotation)
                    .FirstOrDefault(fi => (fi as FamilyInstance)?.Symbol?.Id == symbol.Id) as FamilyInstance;

                if (sampleInstance != null)
                {
                    foreach (Parameter p in sampleInstance.Parameters)
                    {
                        if (p.StorageType == StorageType.String && !p.IsReadOnly
                            && paramItems.All(pi => pi.Name != p.Definition.Name))
                            paramItems.Add(new ParameterItem { Name = p.Definition.Name, ParamDefinition = p.Definition });
                    }
                }
            }

            var sortedParams = paramItems.OrderBy(p => p.Name).ToList();
            CmbParameters.ItemsSource = sortedParams;
            // Restore previously selected parameter if available
            int restoredParamIdx = _pendingParameterRestore != null
                ? sortedParams.FindIndex(p => p.Name == _pendingParameterRestore)
                : -1;
            CmbParameters.SelectedIndex = restoredParamIdx >= 0 ? restoredParamIdx : (sortedParams.Count > 0 ? 0 : -1);
            _pendingParameterRestore = null;
        }

        // ── Grid data collection ──

        /// <summary>
        /// Collect grid data from the view. If filterIds is provided, only include grids with those IDs.
        /// </summary>
        public static List<GridRowItem> CollectGridData(Document doc, View view, HashSet<ElementId> filterIds = null)
        {
            var grids = new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(RevitGrid))
                .Cast<RevitGrid>()
                .Where(g => filterIds == null || filterIds.Contains(g.Id))
                .ToList();

            var viewRight = view.RightDirection;
            var items = new List<GridRowItem>();

            foreach (var grid in grids)
            {
                var curve = grid.Curve;
                string orientation;

                if (curve is Line line)
                {
                    var dir = line.Direction.Normalize();
                    double dot = Math.Abs(dir.DotProduct(viewRight));

                    if (dot >= 0.985)       // within ~10° of horizontal
                        orientation = "Horizontal";
                    else if (dot <= 0.174)   // within ~10° of vertical
                        orientation = "Vertical";
                    else
                        orientation = "Angled";
                }
                else
                {
                    orientation = "Curved";
                }

                bool defaultSelected = orientation == "Horizontal" || orientation == "Vertical";

                items.Add(new GridRowItem
                {
                    IsSelected = defaultSelected,
                    GridName = grid.Name,
                    Orientation = orientation,
                    DetectedOrientation = orientation,
                    GridId = grid.Id
                });
            }

            items.Sort((a, b) => Utils.NaturalSortCompare(a.GridName, b.GridName));
            return items;
        }

        private void RefreshGridList()
        {
            // Capture scope mode on UI thread before raising event
            bool scopeAll = RbScopeAllVisible?.IsChecked == true;
            bool scopeSelected = RbScopeSelected?.IsChecked == true;
            bool scopePick = RbScopePick?.IsChecked == true;
            var pickedIds = new HashSet<ElementId>(_pickedGridIds);

            _handler.HandlerAction = app =>
            {
                var uidoc = app.ActiveUIDocument;
                if (uidoc == null) return;
                var doc = uidoc.Document;
                var view = uidoc.ActiveView;

                // Determine filter based on scope
                HashSet<ElementId> filterIds = null;
                if (scopeSelected)
                {
                    // Only grids in the current Revit selection
                    var selIds = uidoc.Selection.GetElementIds();
                    var gridIds = new HashSet<ElementId>();
                    foreach (var id in selIds)
                    {
                        var elem = doc.GetElement(id);
                        if (elem is RevitGrid) gridIds.Add(id);
                    }
                    filterIds = gridIds;
                }
                else if (scopePick)
                {
                    filterIds = pickedIds.Count > 0 ? pickedIds : null;
                }
                // scopeAll → filterIds stays null (no filter)

                var newItems = CollectGridData(doc, view, filterIds);
                var viewName = view.Name;
                var viewId = view.Id;

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _lastViewId = viewId;
                    DistributeGrids(newItems);
                    TblViewName.Text = $"View: {viewName}";
                    UpdateExpanderHeaders();
                    UpdateIntersectionCount();
                    UpdateFormatPreview();
                    PopulateDropdowns(doc);
                    SetStatus("Ready.");
                }));
            };
            _externalEvent.Raise();
        }

        // ── Intersection counting ──

        private void UpdateIntersectionCount()
        {
            if (TblIntersectionCount == null) return;

            int hCount = _hGridItems.Count(g => g.IsSelected);
            int vCount = _vGridItems.Count(g => g.IsSelected);
            // Also count any Angled/Curved grids reassigned to H or V
            hCount += _otherGridItems.Count(g => g.IsSelected && g.Orientation == "Horizontal");
            vCount += _otherGridItems.Count(g => g.IsSelected && g.Orientation == "Vertical");
            int total = hCount * vCount;
            TblIntersectionCount.Text = $"{total} intersections  ({hCount} H  x  {vCount} V)";
        }

        // ── Format preview ──

        private void UpdateFormatPreview()
        {
            if (TblFormatPreview == null || TxtLabelTemplate == null) return;

            string template = TxtLabelTemplate.Text ?? "({H},{V})";
            string hSample = "A";
            string vSample = "1";

            // Try to use actual grid names if available
            var hGrid = _hGridItems.FirstOrDefault(g => g.IsSelected);
            var vGrid = _vGridItems.FirstOrDefault(g => g.IsSelected);
            if (hGrid != null) hSample = hGrid.GridName;
            if (vGrid != null) vSample = vGrid.GridName;

            string result;
            if (RbOrderHV?.IsChecked == true)
                result = template.Replace("{H}", hSample).Replace("{V}", vSample);
            else
                result = template.Replace("{H}", vSample).Replace("{V}", hSample);

            TblFormatPreview.Text = result;
        }

        // ── Extensible Storage helpers ──

        private static Schema GetOrCreateSchema()
        {
            var schema = Schema.Lookup(SchemaGuid);
            if (schema != null) return schema;

            var builder = new SchemaBuilder(SchemaGuid);
            builder.SetSchemaName(SchemaName);
            builder.SetReadAccessLevel(AccessLevel.Public);
            builder.SetWriteAccessLevel(AccessLevel.Public);
            builder.AddSimpleField(FieldViewId, typeof(string));
            builder.AddSimpleField(FieldHGridName, typeof(string));
            builder.AddSimpleField(FieldVGridName, typeof(string));
            return builder.Finish();
        }

        private static void TagElement(Element element, string viewIdStr, string hName, string vName)
        {
            if (element == null) return;

            var schema = GetOrCreateSchema();
            var entity = new Entity(schema);
            entity.Set<string>(FieldViewId, viewIdStr);
            entity.Set<string>(FieldHGridName, hName);
            entity.Set<string>(FieldVGridName, vName);
            element.SetEntity(entity);
        }

        private static string ViewIdToString(ElementId id)
        {
            return id.ToString();
        }

        private static List<Element> FindTaggedElements(Document doc, View view)
        {
            var schema = Schema.Lookup(SchemaGuid);
            if (schema == null) return new List<Element>();

            string viewIdStr = ViewIdToString(view.Id);

            var candidates = new List<Element>();

            var textNotes = new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(TextNote))
                .ToList();
            candidates.AddRange(textNotes);

            var annotations = new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_GenericAnnotation)
                .ToList();
            candidates.AddRange(annotations);

            var tagged = new List<Element>();
            foreach (var elem in candidates)
            {
                var entity = elem.GetEntity(schema);
                if (entity.IsValid() && entity.Get<string>(FieldViewId) == viewIdStr)
                    tagged.Add(elem);
            }

            return tagged;
        }

        private static HashSet<string> GetExistingIntersectionKeys(List<Element> taggedElements)
        {
            var schema = Schema.Lookup(SchemaGuid);
            if (schema == null) return new HashSet<string>();

            var keys = new HashSet<string>();
            foreach (var elem in taggedElements)
            {
                var entity = elem.GetEntity(schema);
                if (entity.IsValid())
                {
                    string h = entity.Get<string>(FieldHGridName);
                    string v = entity.Get<string>(FieldVGridName);
                    keys.Add($"{h}|{v}");
                }
            }
            return keys;
        }

        // ── Core placement logic ──

        private void ExecutePlaceLabels(UIApplication app)
        {
            var doc = app.ActiveUIDocument.Document;
            var view = app.ActiveUIDocument.ActiveView;

            // Validate view type
            if (!(view is ViewPlan))
            {
                Dispatcher.BeginInvoke(new Action(() => SetStatus("Error: active view must be a plan view.")));
                return;
            }

            // Collect settings from UI thread
            bool useTextNote = false;
            TextNoteType selectedTextType = null;
            FamilySymbol selectedSymbol = null;
            ParameterItem selectedParam = null;
            string labelTemplate = "";
            bool orderHV = true;
            bool autoOffset = true;
            double offsetX = 0, offsetY = 0;
            bool deleteExisting = false;
            bool skipExisting = false;
            List<GridRowItem> hGridItems = null, vGridItems = null;

            Dispatcher.Invoke(() =>
            {
                useTextNote = RbTextNote.IsChecked == true;
                selectedTextType = CmbTextNoteTypes.SelectedItem as TextNoteType;
                selectedSymbol = CmbFamilyTypes.SelectedItem as FamilySymbol;
                selectedParam = CmbParameters.SelectedItem as ParameterItem;
                labelTemplate = string.IsNullOrWhiteSpace(TxtLabelTemplate.Text) ? "({H},{V})" : TxtLabelTemplate.Text;
                orderHV = RbOrderHV.IsChecked == true;
                autoOffset = CkOffsetAuto.IsChecked == true;
                double.TryParse(TxtOffsetX.Text, out offsetX);
                double.TryParse(TxtOffsetY.Text, out offsetY);
                deleteExisting = RbDeleteExisting.IsChecked == true;
                skipExisting = RbSkipExisting.IsChecked == true;

                // Gather H grids: main H list + any Other grids reassigned to Horizontal
                hGridItems = _hGridItems.Where(g => g.IsSelected)
                    .Concat(_otherGridItems.Where(g => g.IsSelected && g.Orientation == "Horizontal"))
                    .ToList();
                // Gather V grids: main V list + any Other grids reassigned to Vertical
                vGridItems = _vGridItems.Where(g => g.IsSelected)
                    .Concat(_otherGridItems.Where(g => g.IsSelected && g.Orientation == "Vertical"))
                    .ToList();
            });

            if (hGridItems.Count == 0 || vGridItems.Count == 0)
            {
                Dispatcher.BeginInvoke(new Action(() =>
                    SetStatus("Error: need at least one Horizontal and one Vertical grid selected.")));
                return;
            }

            if (useTextNote && selectedTextType == null)
            {
                Dispatcher.BeginInvoke(new Action(() => SetStatus("Error: no TextNoteType selected.")));
                return;
            }

            if (!useTextNote && (selectedSymbol == null || selectedParam == null))
            {
                Dispatcher.BeginInvoke(new Action(() =>
                    SetStatus("Error: select a family type and parameter.")));
                return;
            }

            // Compute model offset: paper inches → model feet
            double modelOffsetX, modelOffsetY;
            if (autoOffset)
            {
                double scale = view.Scale;
                modelOffsetX = 0.125 * scale / 12.0;
                modelOffsetY = 0.125 * scale / 12.0;
            }
            else
            {
                double scale = view.Scale;
                modelOffsetX = offsetX * scale / 12.0;
                modelOffsetY = offsetY * scale / 12.0;
            }

            int placedCount = 0;
            int deletedCount = 0;
            int skippedCount = 0;
            int errorCount = 0;
            string errorDetails = "";

            using (Transaction tx = new Transaction(doc, "GridCoords: Place Labels"))
            {
                tx.Start();
                try
                {
                    // Handle existing labels
                    var existingElements = FindTaggedElements(doc, view);
                    HashSet<string> existingKeys = new HashSet<string>();

                    if (deleteExisting)
                    {
                        deletedCount = existingElements.Count;
                        foreach (var elem in existingElements)
                            doc.Delete(elem.Id);
                    }
                    else if (skipExisting)
                    {
                        existingKeys = GetExistingIntersectionKeys(existingElements);
                    }

                    // Activate family symbol if needed
                    if (!useTextNote && !selectedSymbol.IsActive)
                    {
                        selectedSymbol.Activate();
                        doc.Regenerate();
                    }

                    // Get crop box for filtering
                    BoundingBoxXYZ cropBox = null;
                    if (view.CropBoxActive)
                        cropBox = view.CropBox;

                    // Resolve grid ElementIds to Revit Grid objects
                    var hGrids = hGridItems.Select(gi => doc.GetElement(gi.GridId) as RevitGrid).Where(g => g != null).ToList();
                    var vGrids = vGridItems.Select(gi => doc.GetElement(gi.GridId) as RevitGrid).Where(g => g != null).ToList();

                    foreach (var hGrid in hGrids)
                    {
                        foreach (var vGrid in vGrids)
                        {
                            try
                            {
                                var hCurve = hGrid.Curve;
                                var vCurve = vGrid.Curve;

                                XYZ intersection;
#pragma warning disable CS0618 // Curve.Intersect deprecated in 2026; old overload still functional
                                IntersectionResultArray results;
                                var compResult = hCurve.Intersect(vCurve, out results);
#pragma warning restore CS0618
                                if (compResult != SetComparisonResult.Overlap || results == null || results.Size == 0)
                                    continue;
                                intersection = results.get_Item(0).XYZPoint;

                                // Filter by crop box
                                if (cropBox != null)
                                {
                                    var min = cropBox.Min;
                                    var max = cropBox.Max;
                                    if (intersection.X < min.X || intersection.X > max.X ||
                                        intersection.Y < min.Y || intersection.Y > max.Y)
                                    {
                                        skippedCount++;
                                        continue;
                                    }
                                }

                                string hName = hGrid.Name;
                                string vName = vGrid.Name;
                                string intersectionKey = $"{hName}|{vName}";

                                // Skip if existing
                                if (skipExisting && existingKeys.Contains(intersectionKey))
                                {
                                    skippedCount++;
                                    continue;
                                }

                                // Build label text
                                string label;
                                if (orderHV)
                                    label = labelTemplate.Replace("{H}", hName).Replace("{V}", vName);
                                else
                                    label = labelTemplate.Replace("{H}", vName).Replace("{V}", hName);

                                // Apply offset
                                var placementPoint = new XYZ(
                                    intersection.X + modelOffsetX,
                                    intersection.Y + modelOffsetY,
                                    intersection.Z);

                                if (useTextNote)
                                {
                                    var textNote = TextNote.Create(doc, view.Id, placementPoint,
                                        label, selectedTextType.Id);
                                    TagElement(textNote, ViewIdToString(view.Id), hName, vName);
                                }
                                else
                                {
                                    var instance = doc.Create.NewFamilyInstance(
                                        placementPoint, selectedSymbol, view);
                                    var param = instance.LookupParameter(selectedParam.Name);
                                    if (param != null && !param.IsReadOnly)
                                        param.Set(label);
                                    TagElement(instance, ViewIdToString(view.Id), hName, vName);
                                }

                                placedCount++;
                            }
                            catch (Exception ex)
                            {
                                errorCount++;
                                errorDetails += $"{hGrid.Name} x {vGrid.Name}: {ex.Message}\n";
                            }
                        }
                    }

                    tx.Commit();
                }
                catch (Exception ex)
                {
                    tx.RollBack();
                    errorCount++;
                    errorDetails = ex.Message;
                }
            }

            // Update UI on WPF thread
            int placed = placedCount;
            int deleted = deletedCount;
            int skipped = skippedCount;
            int errors = errorCount;
            string details = errorDetails;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                ShowResults(placed, deleted, skipped, errors, details);
                SetStatus(errors > 0
                    ? $"{placed} placed, {errors} error(s)."
                    : $"{placed} labels placed.");
            }));
        }

        private void ExecuteDeleteLabels(UIApplication app)
        {
            var doc = app.ActiveUIDocument.Document;
            var view = app.ActiveUIDocument.ActiveView;

            // Capture which grid names are currently in the form
            HashSet<string> hNames = null;
            HashSet<string> vNames = null;
            bool scopeAll = true;

            Dispatcher.Invoke(() =>
            {
                scopeAll = RbScopeAllVisible?.IsChecked == true;
                if (!scopeAll)
                {
                    // Only delete labels whose grid names match what's shown in the form
                    hNames = new HashSet<string>(
                        _hGridItems.Select(g => g.GridName)
                        .Concat(_otherGridItems.Where(g => g.Orientation == "Horizontal").Select(g => g.GridName)));
                    vNames = new HashSet<string>(
                        _vGridItems.Select(g => g.GridName)
                        .Concat(_otherGridItems.Where(g => g.Orientation == "Vertical").Select(g => g.GridName)));
                }
            });

            int deletedCount = 0;

            using (Transaction tx = new Transaction(doc, "GridCoords: Delete Labels"))
            {
                tx.Start();
                try
                {
                    var tagged = FindTaggedElements(doc, view);

                    // Filter to only labels matching the current scope's grids
                    if (!scopeAll && hNames != null && vNames != null)
                    {
                        var schema = Schema.Lookup(SchemaGuid);
                        tagged = tagged.Where(elem =>
                        {
                            var entity = elem.GetEntity(schema);
                            if (!entity.IsValid()) return false;
                            string h = entity.Get<string>(FieldHGridName);
                            string v = entity.Get<string>(FieldVGridName);
                            return hNames.Contains(h) && vNames.Contains(v);
                        }).ToList();
                    }

                    deletedCount = tagged.Count;
                    foreach (var elem in tagged)
                        doc.Delete(elem.Id);
                    tx.Commit();
                }
                catch (Exception ex)
                {
                    tx.RollBack();
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        ShowResults(0, 0, 0, 1, ex.Message);
                        SetStatus($"Error: {ex.Message}");
                    }));
                    return;
                }
            }

            int count = deletedCount;
            Dispatcher.BeginInvoke(new Action(() =>
            {
                ShowResults(0, count, 0, 0, "");
                SetStatus($"{count} labels deleted.");
            }));
        }

        // ── UI helpers ──

        private void SetStatus(string text)
        {
            TblStatus.Text = text;
        }

        private void ShowResults(int placed, int deleted, int skipped, int errors, string details)
        {
            ResultsExpander.Visibility = WinVisibility.Visible;
            ResultsExpander.IsExpanded = true;
            TblResultPlaced.Text = placed.ToString();
            TblResultDeleted.Text = deleted.ToString();
            TblResultSkipped.Text = skipped.ToString();
            TblResultErrors.Text = errors.ToString();

            if (errors > 0 && !string.IsNullOrWhiteSpace(details))
            {
                TblResultDetails.Text = details.TrimEnd();
                TblResultDetails.Visibility = WinVisibility.Visible;
            }
            else
            {
                TblResultDetails.Visibility = WinVisibility.Collapsed;
            }
        }

        // ── Event handlers ──

        private void PlacementMode_Changed(object sender, RoutedEventArgs e)
        {
            if (PnlTextNoteOptions == null || PnlFamilyOptions == null) return;

            if (RbTextNote.IsChecked == true)
            {
                PnlTextNoteOptions.Visibility = WinVisibility.Visible;
                PnlFamilyOptions.Visibility = WinVisibility.Collapsed;
            }
            else
            {
                PnlTextNoteOptions.Visibility = WinVisibility.Collapsed;
                PnlFamilyOptions.Visibility = WinVisibility.Visible;
            }
        }

        private void CmbFamilyTypes_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var symbol = CmbFamilyTypes.SelectedItem as FamilySymbol;
            if (symbol == null) return;

            _handler.HandlerAction = app =>
            {
                var doc = app.ActiveUIDocument.Document;
                Dispatcher.BeginInvoke(new Action(() => PopulateParametersForFamily(doc, symbol)));
            };
            _externalEvent.Raise();
        }

        private void TxtLabelTemplate_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdateFormatPreview();
        }

        private void NameOrder_Changed(object sender, RoutedEventArgs e)
        {
            UpdateFormatPreview();
        }

        private void OffsetAuto_Changed(object sender, RoutedEventArgs e)
        {
            if (TxtOffsetX == null || TxtOffsetY == null) return;
            bool auto = CkOffsetAuto.IsChecked == true;
            TxtOffsetX.IsEnabled = !auto;
            TxtOffsetY.IsEnabled = !auto;
        }

        // ── Grid checkbox events ──

        private void GridCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            if (_suppressSelectAllEvents) return;
            UpdateExpanderHeaders();
            UpdateIntersectionCount();
            UpdateFormatPreview();
        }

        private void CkSelectAllH_Changed(object sender, RoutedEventArgs e)
        {
            if (CkSelectAllH == null) return;
            _suppressSelectAllEvents = true;
            bool check = CkSelectAllH.IsChecked == true;
            foreach (var item in _hGridItems) item.IsSelected = check;
            _suppressSelectAllEvents = false;
            UpdateExpanderHeaders();
            UpdateIntersectionCount();
            UpdateFormatPreview();
        }

        private void CkSelectAllV_Changed(object sender, RoutedEventArgs e)
        {
            if (CkSelectAllV == null) return;
            _suppressSelectAllEvents = true;
            bool check = CkSelectAllV.IsChecked == true;
            foreach (var item in _vGridItems) item.IsSelected = check;
            _suppressSelectAllEvents = false;
            UpdateExpanderHeaders();
            UpdateIntersectionCount();
            UpdateFormatPreview();
        }

        private void OtherGridOrientation_Changed(object sender, SelectionChangedEventArgs e)
        {
            // When user reassigns an Angled/Curved grid to H or V, update counts
            UpdateExpanderHeaders();
            UpdateIntersectionCount();
            UpdateFormatPreview();
        }

        // ── Scope event handlers ──

        private void GridScope_Changed(object sender, RoutedEventArgs e)
        {
            if (BtnPick == null || BtnClearPick == null) return;

            bool pickMode = RbScopePick.IsChecked == true;
            BtnPick.IsEnabled = pickMode;
            BtnClearPick.IsEnabled = pickMode;

            // Refresh the grid list to apply the new scope
            RefreshGridList();
        }

        private void BtnPick_Click(object sender, RoutedEventArgs e)
        {
            // Minimize form, let user pick grids, then restore
            this.WindowState = WindowState.Minimized;

            _handler.HandlerAction = app =>
            {
                var uidoc = app.ActiveUIDocument;
                if (uidoc == null) return;

                try
                {
                    var refs = uidoc.Selection.PickObjects(
                        ObjectType.Element,
                        new GridSelectionFilter(),
                        "Select grids, then press Finish or Escape.");

                    _pickedGridIds = refs.Select(r => r.ElementId).ToList();
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    // User pressed Escape — keep previous picks
                }

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    this.WindowState = WindowState.Normal;
                    this.Activate();

                    // Update pick count label
                    if (_pickedGridIds.Count > 0)
                    {
                        TblPickCount.Text = $"{_pickedGridIds.Count} picked";
                        TblPickCount.Visibility = WinVisibility.Visible;
                    }
                    else
                    {
                        TblPickCount.Visibility = WinVisibility.Collapsed;
                    }

                    // Refresh with picked grids
                    RefreshGridList();
                }));
            };
            _externalEvent.Raise();
        }

        private void BtnClearPick_Click(object sender, RoutedEventArgs e)
        {
            _pickedGridIds.Clear();
            TblPickCount.Visibility = WinVisibility.Collapsed;
            RefreshGridList();
        }

        private void BtnRefreshGrids_Click(object sender, RoutedEventArgs e)
        {
            SetStatus("Refreshing grids...");
            RefreshGridList();
        }

        private void BtnPlaceLabels_Click(object sender, RoutedEventArgs e)
        {
            SetStatus("Placing labels...");
            BtnPlaceLabels.IsEnabled = false;
            BtnDeleteLabels.IsEnabled = false;

            _handler.HandlerAction = app =>
            {
                try
                {
                    ExecutePlaceLabels(app);
                }
                finally
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        BtnPlaceLabels.IsEnabled = true;
                        BtnDeleteLabels.IsEnabled = true;
                    }));
                }
            };
            _externalEvent.Raise();
        }

        private void BtnDeleteLabels_Click(object sender, RoutedEventArgs e)
        {
            SetStatus("Deleting labels...");
            BtnPlaceLabels.IsEnabled = false;
            BtnDeleteLabels.IsEnabled = false;

            _handler.HandlerAction = app =>
            {
                try
                {
                    ExecuteDeleteLabels(app);
                }
                finally
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        BtnPlaceLabels.IsEnabled = true;
                        BtnDeleteLabels.IsEnabled = true;
                    }));
                }
            };
            _externalEvent.Raise();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }

    // ── Data models ──

    public class GridRowItem : INotifyPropertyChanged
    {
        private bool _isSelected;
        private string _orientation;

        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; OnPropertyChanged(); }
        }

        public string GridName { get; set; }

        public string Orientation
        {
            get => _orientation;
            set { _orientation = value; OnPropertyChanged(); }
        }

        public string DetectedOrientation { get; set; }
        public ElementId GridId { get; set; }

        /// <summary>Options shown in the Angled/Curved section for reassignment.</summary>
        public List<string> AssignableOrientations { get; } =
            new List<string> { "Angled", "Curved", "Horizontal", "Vertical" };

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class ParameterItem
    {
        public string Name { get; set; }
        public Definition ParamDefinition { get; set; }
    }

    /// <summary>
    /// Selection filter that only allows Grid elements to be picked.
    /// </summary>
    public class GridSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem) => elem is RevitGrid;
        public bool AllowReference(Reference reference, XYZ position) => false;
    }
}
