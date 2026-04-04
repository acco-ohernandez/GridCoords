using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
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

        // ── State: observable collections for H and V grids (bound to XAML-defined ItemsControls) ──
        private readonly ObservableCollection<GridRowItem> _hGridItems = new ObservableCollection<GridRowItem>();
        private readonly ObservableCollection<GridRowItem> _vGridItems = new ObservableCollection<GridRowItem>();
        private ElementId _lastViewId;
        private bool _suppressSelectAllEvents;

        // ── Angle-based grid groups and intersection pairings ──
        private readonly List<GridGroup> _allGridGroups = new List<GridGroup>();
        private readonly ObservableCollection<IntersectionPairing> _intersectionPairings = new ObservableCollection<IntersectionPairing>();

        /// <summary>
        /// Tolerance in degrees for clustering grids into the same angle group.
        /// Grids within this angular distance of each other are grouped together.
        /// </summary>
        private const double AngleClusteringToleranceDegrees = 10.0;

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

            // Bind the XAML-defined ItemsControls for standard H/V groups
            HGridList.ItemsSource = _hGridItems;
            VGridList.ItemsSource = _vGridItems;

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
            _allGridGroups.Clear();
            _intersectionPairings.Clear();

            // Cluster grids by angle into groups
            var clusteredGroups = ClusterGridsIntoGroups(allGrids);
            _allGridGroups.AddRange(clusteredGroups);

            // Populate the XAML-bound standard collections (H/V groups use XAML-defined expanders)
            foreach (var gridGroup in clusteredGroups)
            {
                if (gridGroup.GroupClassification == "Horizontal")
                {
                    foreach (var gridItem in gridGroup.GridItems)
                        _hGridItems.Add(gridItem);
                }
                else if (gridGroup.GroupClassification == "Vertical")
                {
                    foreach (var gridItem in gridGroup.GridItems)
                        _vGridItems.Add(gridItem);
                }
            }

            // Generate intersection pairings from all groups
            var generatedPairings = GenerateIntersectionPairings(clusteredGroups);
            foreach (var pairing in generatedPairings)
                _intersectionPairings.Add(pairing);

            // Build dynamic UI for non-standard groups (diagonal, curved)
            BuildDynamicGroupExpanders();
            BuildPairingsUI();
        }

        /// <summary>
        /// Programmatically creates expander sections for non-standard grid groups (diagonal, curved).
        /// Each expander mirrors the visual structure of the XAML-defined H/V expanders.
        /// </summary>
        private void BuildDynamicGroupExpanders()
        {
            if (DynamicGroupsContainer == null || ExpanderIntersectionPairings == null) return;
            DynamicGroupsContainer.Children.Clear();

            var nonStandardGroups = _allGridGroups.Where(group => !group.IsStandardGroup).ToList();
            if (nonStandardGroups.Count == 0) return;

            foreach (var gridGroup in nonStandardGroups)
            {
                // Create the expander matching the XAML Expander style
                var groupExpander = new Expander
                {
                    Header = $"{gridGroup.GroupDisplayLabel} ({gridGroup.SelectedGridCount} of {gridGroup.GridItems.Count})",
                    IsExpanded = true,
                    Margin = new Thickness(0, 0, 0, 6),
                    Padding = new Thickness(6, 4, 6, 4),
                    FontWeight = FontWeights.SemiBold,
                    BorderBrush = new System.Windows.Media.SolidColorBrush(
                        (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#CCCCCC")),
                    BorderThickness = new Thickness(1)
                };

                // Store reference to the group for event handling
                groupExpander.Tag = gridGroup;

                var contentDockPanel = new DockPanel { Margin = new Thickness(2, 4, 2, 4) };

                // Select All checkbox at the top
                var selectAllCheckBox = new System.Windows.Controls.CheckBox
                {
                    Content = "Select All",
                    IsChecked = true,
                    FontWeight = FontWeights.Normal,
                    Margin = new Thickness(0, 2, 0, 2)
                };
                selectAllCheckBox.Tag = gridGroup;
                selectAllCheckBox.Checked += DynamicSelectAll_Changed;
                selectAllCheckBox.Unchecked += DynamicSelectAll_Changed;

                var selectAllDockPanel = new DockPanel { Margin = new Thickness(0, 0, 0, 4) };
                selectAllDockPanel.Children.Add(selectAllCheckBox);
                DockPanel.SetDock(selectAllDockPanel, Dock.Top);
                contentDockPanel.Children.Add(selectAllDockPanel);

                // Scrollable WrapPanel of grid checkboxes
                var gridItemsWrapPanel = new System.Windows.Controls.WrapPanel
                {
                    Orientation = System.Windows.Controls.Orientation.Horizontal
                };

                foreach (var gridItem in gridGroup.GridItems)
                {
                    var gridCheckBox = new System.Windows.Controls.CheckBox
                    {
                        Content = gridItem.GridName,
                        FontWeight = FontWeights.Normal,
                        Margin = new Thickness(0, 2, 12, 2),
                        Tag = gridGroup
                    };

                    // Bind IsChecked to the grid item
                    var isCheckedBinding = new System.Windows.Data.Binding("IsSelected")
                    {
                        Source = gridItem,
                        Mode = System.Windows.Data.BindingMode.TwoWay
                    };
                    gridCheckBox.SetBinding(System.Windows.Controls.Primitives.ToggleButton.IsCheckedProperty, isCheckedBinding);

                    gridCheckBox.Checked += DynamicGridCheckBox_Changed;
                    gridCheckBox.Unchecked += DynamicGridCheckBox_Changed;
                    gridItemsWrapPanel.Children.Add(gridCheckBox);
                }

                var gridScrollViewer = new ScrollViewer
                {
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto
                };
                gridScrollViewer.Content = gridItemsWrapPanel;
                contentDockPanel.Children.Add(gridScrollViewer);

                groupExpander.Content = contentDockPanel;
                DynamicGroupsContainer.Children.Add(groupExpander);
            }
        }

        /// <summary>
        /// Builds the Intersection Pairings UI section showing cross-group pairings
        /// with enable/disable checkboxes and per-pairing intersection counts.
        /// </summary>
        private void BuildPairingsUI()
        {
            if (PairingsContainer == null || ExpanderIntersectionPairings == null) return;
            PairingsContainer.Children.Clear();

            // Only show pairings section when there are multiple groups to pair
            if (_intersectionPairings.Count <= 1)
            {
                ExpanderIntersectionPairings.Visibility = WinVisibility.Visible;
                // Still show the single standard pairing for clarity
            }

            if (_intersectionPairings.Count == 0)
            {
                ExpanderIntersectionPairings.Visibility = WinVisibility.Collapsed;
                return;
            }

            ExpanderIntersectionPairings.Visibility = WinVisibility.Visible;

            foreach (var pairing in _intersectionPairings)
            {
                var pairingDockPanel = new DockPanel { Margin = new Thickness(0, 2, 0, 2) };

                var enabledCheckBox = new System.Windows.Controls.CheckBox
                {
                    FontWeight = FontWeights.Normal,
                    Margin = new Thickness(0, 0, 6, 0),
                    Tag = pairing
                };

                // Bind IsChecked to pairing.IsEnabled
                var isEnabledBinding = new System.Windows.Data.Binding("IsEnabled")
                {
                    Source = pairing,
                    Mode = System.Windows.Data.BindingMode.TwoWay
                };
                enabledCheckBox.SetBinding(System.Windows.Controls.Primitives.ToggleButton.IsCheckedProperty, isEnabledBinding);
                enabledCheckBox.Checked += PairingCheckBox_Changed;
                enabledCheckBox.Unchecked += PairingCheckBox_Changed;

                var pairingLabelTextBlock = new TextBlock
                {
                    FontWeight = FontWeights.Normal,
                    VerticalAlignment = VerticalAlignment.Center
                };

                // Build descriptive label with H/V token assignment
                string horizontalGroupLabel = pairing.GroupForHToken.GroupClassification;
                string verticalGroupLabel = pairing.GroupForVToken.GroupClassification;
                int horizontalSelectedCount = pairing.GroupForHToken.SelectedGridCount;
                int verticalSelectedCount = pairing.GroupForVToken.SelectedGridCount;
                int estimatedCount = horizontalSelectedCount * verticalSelectedCount;
                pairingLabelTextBlock.Text = $"{horizontalGroupLabel} {{H}} × {verticalGroupLabel} {{V}}  =  {estimatedCount}";

                pairingDockPanel.Children.Add(enabledCheckBox);
                pairingDockPanel.Children.Add(pairingLabelTextBlock);
                PairingsContainer.Children.Add(pairingDockPanel);
            }
        }

        /// <summary>
        /// Event handler for Select All checkbox on dynamically generated group expanders.
        /// </summary>
        private void DynamicSelectAll_Changed(object sender, RoutedEventArgs eventArgs)
        {
            var selectAllCheckBox = sender as System.Windows.Controls.CheckBox;
            var gridGroup = selectAllCheckBox?.Tag as GridGroup;
            if (gridGroup == null || gridGroup.IsSuppressingEvents) return;

            bool isChecked = selectAllCheckBox.IsChecked == true;
            gridGroup.ApplySelectAllToItems(isChecked);
            UpdateExpanderHeaders();
            UpdateIntersectionCount();
            UpdateFormatPreview();
        }

        /// <summary>
        /// Event handler for individual grid checkboxes in dynamically generated group expanders.
        /// </summary>
        private void DynamicGridCheckBox_Changed(object sender, RoutedEventArgs eventArgs)
        {
            var gridCheckBox = sender as System.Windows.Controls.CheckBox;
            var gridGroup = gridCheckBox?.Tag as GridGroup;
            if (gridGroup == null || gridGroup.IsSuppressingEvents) return;

            UpdateExpanderHeaders();
            UpdateIntersectionCount();
            UpdateFormatPreview();
        }

        /// <summary>
        /// Event handler for pairing enable/disable checkboxes in the Intersection Pairings section.
        /// </summary>
        private void PairingCheckBox_Changed(object sender, RoutedEventArgs eventArgs)
        {
            UpdateIntersectionCount();
        }

        private void UpdateExpanderHeaders()
        {
            if (ExpanderHGrids == null || ExpanderVGrids == null) return;

            int horizontalSelectedCount = _hGridItems.Count(gridItem => gridItem.IsSelected);
            int verticalSelectedCount = _vGridItems.Count(gridItem => gridItem.IsSelected);
            ExpanderHGrids.Header = $"Horizontal Grids ({horizontalSelectedCount} of {_hGridItems.Count})";
            ExpanderVGrids.Header = $"Vertical Grids ({verticalSelectedCount} of {_vGridItems.Count})";

            // Update headers for dynamically generated group expanders
            if (DynamicGroupsContainer != null)
            {
                foreach (var child in DynamicGroupsContainer.Children)
                {
                    if (child is Expander dynamicExpander && dynamicExpander.Tag is GridGroup gridGroup)
                    {
                        dynamicExpander.Header = $"{gridGroup.GroupDisplayLabel} ({gridGroup.SelectedGridCount} of {gridGroup.GridItems.Count})";
                    }
                }
            }
        }

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
            catch (Exception ex)
            {
                // Revit may not be ready during startup or shutdown — log but don't surface
                Debug.WriteLine($"GridCoords ThreadIdle: {ex.Message}");
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

            var viewRightDirection = view.RightDirection;
            var gridRowItems = new List<GridRowItem>();

            foreach (var grid in grids)
            {
                var gridCurve = grid.Curve;
                string orientation;
                double angleFromHorizontalDegrees = -1;

                if (gridCurve is Line gridLine)
                {
                    var lineDirection = gridLine.Direction.Normalize();
                    double dotProductWithRight = lineDirection.DotProduct(viewRightDirection);
                    double absoluteDotProduct = Math.Abs(dotProductWithRight);

                    // Compute angle in degrees (0–180°) relative to view horizontal.
                    // acos(|dot|) gives 0° when parallel to horizontal, 90° when perpendicular.
                    // Normalize to 0–180° range so parallel grids always cluster together.
                    double angleRadians = Math.Acos(Math.Min(absoluteDotProduct, 1.0));
                    angleFromHorizontalDegrees = angleRadians * 180.0 / Math.PI;

                    if (absoluteDotProduct >= 0.985)       // within ~10° of horizontal
                        orientation = "Horizontal";
                    else if (absoluteDotProduct <= 0.174)   // within ~10° of vertical
                        orientation = "Vertical";
                    else
                        orientation = "Angled";
                }
                else
                {
                    orientation = "Curved";
                }

                bool isSelectedByDefault = orientation == "Horizontal" || orientation == "Vertical";

                gridRowItems.Add(new GridRowItem
                {
                    IsSelected = isSelectedByDefault,
                    GridName = grid.Name,
                    Orientation = orientation,
                    AngleFromViewHorizontal = angleFromHorizontalDegrees,
                    GridId = grid.Id
                });
            }

            gridRowItems.Sort((a, b) => Utils.NaturalSortCompare(a.GridName, b.GridName));
            return gridRowItems;
        }

        /// <summary>
        /// Clusters grid items into angle-based groups using circular distance.
        /// Groups within AngleClusteringToleranceDegrees are merged.
        /// Standard groups (Horizontal ~0°, Vertical ~90°) are always created even if empty.
        /// </summary>
        public static List<GridGroup> ClusterGridsIntoGroups(List<GridRowItem> allGridRowItems)
        {
            // Separate curved grids (no meaningful angle) from line grids
            var lineGridItems = allGridRowItems.Where(gridItem => gridItem.AngleFromViewHorizontal >= 0).ToList();
            var curvedGridItems = allGridRowItems.Where(gridItem => gridItem.AngleFromViewHorizontal < 0).ToList();

            // Sort line grids by angle for clustering
            var sortedByAngle = lineGridItems.OrderBy(gridItem => gridItem.AngleFromViewHorizontal).ToList();

            // Greedy clustering: walk sorted angles, start a new cluster when the gap exceeds tolerance
            var angleClusters = new List<List<GridRowItem>>();
            List<GridRowItem> currentCluster = null;
            double currentClusterStartAngle = -999;

            foreach (var gridItem in sortedByAngle)
            {
                double circularDistance = currentCluster == null
                    ? double.MaxValue
                    : ComputeCircularAngleDistance(gridItem.AngleFromViewHorizontal, currentClusterStartAngle);

                if (currentCluster == null || circularDistance > AngleClusteringToleranceDegrees)
                {
                    // Start a new cluster
                    currentCluster = new List<GridRowItem> { gridItem };
                    angleClusters.Add(currentCluster);
                    currentClusterStartAngle = gridItem.AngleFromViewHorizontal;
                }
                else
                {
                    currentCluster.Add(gridItem);
                }
            }

            // Also check if the first and last clusters should merge (wraparound at 0°/180° boundary)
            if (angleClusters.Count >= 2)
            {
                var firstCluster = angleClusters[0];
                var lastCluster = angleClusters[angleClusters.Count - 1];
                double firstClusterAngle = firstCluster[0].AngleFromViewHorizontal;
                double lastClusterAngle = lastCluster[lastCluster.Count - 1].AngleFromViewHorizontal;

                if (ComputeCircularAngleDistance(firstClusterAngle, lastClusterAngle) <= AngleClusteringToleranceDegrees)
                {
                    // Merge last cluster into first
                    foreach (var gridItem in lastCluster)
                        firstCluster.Add(gridItem);
                    angleClusters.RemoveAt(angleClusters.Count - 1);
                }
            }

            // Convert clusters into GridGroup objects
            var gridGroups = new List<GridGroup>();

            foreach (var cluster in angleClusters)
            {
                double averageAngle = cluster.Average(gridItem => gridItem.AngleFromViewHorizontal);
                string classification;
                string displayLabel;
                bool isStandard;

                if (averageAngle <= AngleClusteringToleranceDegrees || averageAngle >= (180 - AngleClusteringToleranceDegrees))
                {
                    classification = "Horizontal";
                    displayLabel = "Horizontal Grids (0°)";
                    isStandard = true;
                    averageAngle = 0;
                }
                else if (Math.Abs(averageAngle - 90) <= AngleClusteringToleranceDegrees)
                {
                    classification = "Vertical";
                    displayLabel = "Vertical Grids (90°)";
                    isStandard = true;
                    averageAngle = 90;
                }
                else
                {
                    classification = "Diagonal";
                    displayLabel = $"Diagonal Grids (~{Math.Round(averageAngle)}°)";
                    isStandard = false;
                }

                var gridGroup = new GridGroup
                {
                    GroupDisplayLabel = displayLabel,
                    RepresentativeAngleDegrees = averageAngle,
                    IsStandardGroup = isStandard,
                    GroupClassification = classification
                };

                foreach (var gridItem in cluster)
                {
                    gridItem.Orientation = classification;
                    gridItem.IsSelected = true;
                    gridGroup.GridItems.Add(gridItem);
                }

                gridGroups.Add(gridGroup);
            }

            // Add curved grids as a separate group (if any exist)
            if (curvedGridItems.Count > 0)
            {
                var curvedGroup = new GridGroup
                {
                    GroupDisplayLabel = $"Curved Grids ({curvedGridItems.Count})",
                    RepresentativeAngleDegrees = -1,
                    IsStandardGroup = false,
                    GroupClassification = "Curved"
                };
                foreach (var gridItem in curvedGridItems)
                    curvedGroup.GridItems.Add(gridItem);
                gridGroups.Add(curvedGroup);
            }

            // Sort groups by angle (Horizontal first, then ascending, Curved last)
            gridGroups.Sort((groupA, groupB) =>
                groupA.RepresentativeAngleDegrees.CompareTo(groupB.RepresentativeAngleDegrees));

            return gridGroups;
        }

        /// <summary>
        /// Computes the circular (modular) distance between two angles in the 0°–180° range.
        /// Handles the wraparound where 5° and 175° are 10° apart, not 170°.
        /// </summary>
        private static double ComputeCircularAngleDistance(double angleDegrees1, double angleDegrees2)
        {
            double absoluteDifference = Math.Abs(angleDegrees1 - angleDegrees2);
            return Math.Min(absoluteDifference, 180.0 - absoluteDifference);
        }

        /// <summary>
        /// Generates all cross-group intersection pairings from the given grid groups.
        /// Assigns {H} token to the group with the smaller angle ("more horizontal").
        /// Excludes curved-only groups and same-group pairings.
        /// </summary>
        private static List<IntersectionPairing> GenerateIntersectionPairings(List<GridGroup> gridGroups)
        {
            var pairings = new List<IntersectionPairing>();
            var pairableGroups = gridGroups
                .Where(group => group.GroupClassification != "Curved" && group.GridItems.Count > 0)
                .ToList();

            for (int outerIndex = 0; outerIndex < pairableGroups.Count; outerIndex++)
            {
                for (int innerIndex = outerIndex + 1; innerIndex < pairableGroups.Count; innerIndex++)
                {
                    var groupWithSmallerAngle = pairableGroups[outerIndex];
                    var groupWithLargerAngle = pairableGroups[innerIndex];

                    // The group with the smaller angle is "more horizontal" and gets the {H} token
                    var pairing = new IntersectionPairing
                    {
                        GroupForHToken = groupWithSmallerAngle,
                        GroupForVToken = groupWithLargerAngle,
                        PairingDisplayLabel = $"{groupWithSmallerAngle.GroupClassification} × {groupWithLargerAngle.GroupClassification}"
                    };

                    pairings.Add(pairing);
                }
            }

            return pairings;
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

            // Sum intersection counts across all enabled pairings
            int totalIntersectionEstimate = 0;
            var pairingBreakdownParts = new List<string>();

            foreach (var pairing in _intersectionPairings)
            {
                if (!pairing.IsEnabled) continue;
                int pairingCount = pairing.EstimatedIntersectionCount;
                totalIntersectionEstimate += pairingCount;
                pairingBreakdownParts.Add(
                    $"{pairing.GroupForHToken.SelectedGridCount}{pairing.GroupForHToken.GroupClassification[0]} × " +
                    $"{pairing.GroupForVToken.SelectedGridCount}{pairing.GroupForVToken.GroupClassification[0]}");
            }

            string breakdownText = pairingBreakdownParts.Count > 0
                ? string.Join(" + ", pairingBreakdownParts)
                : "no pairings enabled";

            TblIntersectionCount.Text = $"{totalIntersectionEstimate} intersections  ({breakdownText})";

            // Also refresh the pairings UI counts
            RefreshPairingsCountLabels();
        }

        /// <summary>
        /// Refreshes the intersection count labels in the pairings UI without rebuilding it.
        /// </summary>
        private void RefreshPairingsCountLabels()
        {
            if (PairingsContainer == null) return;

            int childIndex = 0;
            foreach (var pairing in _intersectionPairings)
            {
                if (childIndex >= PairingsContainer.Children.Count) break;
                var pairingDockPanel = PairingsContainer.Children[childIndex] as DockPanel;
                if (pairingDockPanel != null && pairingDockPanel.Children.Count >= 2)
                {
                    var labelTextBlock = pairingDockPanel.Children[1] as TextBlock;
                    if (labelTextBlock != null)
                    {
                        int horizontalSelectedCount = pairing.GroupForHToken.SelectedGridCount;
                        int verticalSelectedCount = pairing.GroupForVToken.SelectedGridCount;
                        int estimatedCount = horizontalSelectedCount * verticalSelectedCount;
                        labelTextBlock.Text = $"{pairing.GroupForHToken.GroupClassification} {{H}} × " +
                                              $"{pairing.GroupForVToken.GroupClassification} {{V}}  =  {estimatedCount}";
                    }
                }
                childIndex++;
            }
        }

        // ── Format preview ──

        private void UpdateFormatPreview()
        {
            if (TblFormatPreview == null || TxtLabelTemplate == null) return;

            string template = TxtLabelTemplate.Text ?? "({H},{V})";
            string horizontalSampleName = "A";
            string verticalSampleName = "1";

            // Use the first enabled pairing's grids for realistic sample names
            var firstEnabledPairing = _intersectionPairings.FirstOrDefault(pairing => pairing.IsEnabled);
            if (firstEnabledPairing != null)
            {
                var horizontalSampleGrid = firstEnabledPairing.GroupForHToken.GridItems
                    .FirstOrDefault(gridItem => gridItem.IsSelected);
                var verticalSampleGrid = firstEnabledPairing.GroupForVToken.GridItems
                    .FirstOrDefault(gridItem => gridItem.IsSelected);
                if (horizontalSampleGrid != null) horizontalSampleName = horizontalSampleGrid.GridName;
                if (verticalSampleGrid != null) verticalSampleName = verticalSampleGrid.GridName;
            }
            else
            {
                // Fallback: try the legacy H/V collections
                var horizontalGrid = _hGridItems.FirstOrDefault(gridItem => gridItem.IsSelected);
                var verticalGrid = _vGridItems.FirstOrDefault(gridItem => gridItem.IsSelected);
                if (horizontalGrid != null) horizontalSampleName = horizontalGrid.GridName;
                if (verticalGrid != null) verticalSampleName = verticalGrid.GridName;
            }

            string previewResult;
            if (RbOrderHV?.IsChecked == true)
                previewResult = template.Replace("{H}", horizontalSampleName).Replace("{V}", verticalSampleName);
            else
                previewResult = template.Replace("{H}", verticalSampleName).Replace("{V}", horizontalSampleName);

            TblFormatPreview.Text = previewResult;
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

        /// <summary>
        /// Gets the 2D tangent direction of a curve, normalized.
        /// For lines, returns the line direction. For curves, returns the tangent at midpoint.
        /// </summary>
        private static XYZ GetCurveTangentDirection(Curve curve)
        {
            if (curve is Line line)
            {
                var direction = line.Direction;
                return new XYZ(direction.X, direction.Y, 0).Normalize();
            }
            else
            {
                // For curved grids, use the tangent at the midpoint
                double midParameter = (curve.GetEndParameter(0) + curve.GetEndParameter(1)) / 2.0;
                var tangent = curve.ComputeDerivatives(midParameter, false).BasisX;
                return new XYZ(tangent.X, tangent.Y, 0).Normalize();
            }
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
            bool alignAboveGrid = true;
            bool alignRightOfGrid = true;
            bool rotateDiagonalLabels = true;

            // Snapshot the enabled pairings and their grid items from the UI thread
            var enabledPairingSnapshots = new List<PairingSnapshot>();

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
                alignAboveGrid = RbAlignAbove.IsChecked == true;
                alignRightOfGrid = RbAlignRight.IsChecked == true;
                rotateDiagonalLabels = CkRotateDiagonal.IsChecked == true;

                // Gather enabled pairings with their selected grid items
                foreach (var pairing in _intersectionPairings)
                {
                    if (!pairing.IsEnabled) continue;
                    var horizontalGridsInPairing = pairing.GroupForHToken.GridItems
                        .Where(gridItem => gridItem.IsSelected).ToList();
                    var verticalGridsInPairing = pairing.GroupForVToken.GridItems
                        .Where(gridItem => gridItem.IsSelected).ToList();
                    if (horizontalGridsInPairing.Count > 0 && verticalGridsInPairing.Count > 0)
                    {
                        enabledPairingSnapshots.Add(new PairingSnapshot
                        {
                            HorizontalGridItems = horizontalGridsInPairing,
                            VerticalGridItems = verticalGridsInPairing,
                            HGroupAngleDegrees = pairing.GroupForHToken.RepresentativeAngleDegrees,
                            VGroupAngleDegrees = pairing.GroupForVToken.RepresentativeAngleDegrees
                        });
                    }
                }
            });

            if (enabledPairingSnapshots.Count == 0)
            {
                Dispatcher.BeginInvoke(new Action(() =>
                    SetStatus("Error: no enabled pairings with selected grids.")));
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

            // Compute model offset values: paper inches → model feet
            double viewScale = view.Scale;
            double manualModelOffsetX = offsetX * viewScale / 12.0;
            double manualModelOffsetY = offsetY * viewScale / 12.0;

            // Auto offset: small padding between the label edge and the grid line
            double paddingModelUnits = (1.0 / 32.0) * viewScale / 12.0;

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

                    // ── Phase 1: Create all elements at intersection points ──
                    // Elements are created and tagged with Extensible Storage, then queued
                    // for post-creation alignment. This avoids calling doc.Regenerate() per
                    // element — instead, one single Regenerate after all creates.
                    var pendingLabelAlignments = new List<PendingLabelAlignment>();

                    foreach (var pairingSnapshot in enabledPairingSnapshots)
                    {
                        var horizontalRevitGrids = pairingSnapshot.HorizontalGridItems
                            .Select(gridItem => doc.GetElement(gridItem.GridId) as RevitGrid)
                            .Where(grid => grid != null).ToList();
                        var verticalRevitGrids = pairingSnapshot.VerticalGridItems
                            .Select(gridItem => doc.GetElement(gridItem.GridId) as RevitGrid)
                            .Where(grid => grid != null).ToList();

                        foreach (var horizontalGrid in horizontalRevitGrids)
                        {
                            foreach (var verticalGrid in verticalRevitGrids)
                            {
                                try
                                {
                                    var horizontalCurve = horizontalGrid.Curve;
                                    var verticalCurve = verticalGrid.Curve;

#pragma warning disable CS0618 // Curve.Intersect deprecated in 2026; old overload still functional
                                    IntersectionResultArray intersectionResults;
                                    var comparisonResult = horizontalCurve.Intersect(verticalCurve, out intersectionResults);
#pragma warning restore CS0618
                                    if (comparisonResult != SetComparisonResult.Overlap || intersectionResults == null || intersectionResults.Size == 0)
                                        continue;

                                    // Handle multiple intersections (e.g., curved grids can intersect at 2+ points)
                                    for (int intersectionIndex = 0; intersectionIndex < intersectionResults.Size; intersectionIndex++)
                                    {
                                        XYZ intersectionPoint = intersectionResults.get_Item(intersectionIndex).XYZPoint;

                                        // Filter by crop box
                                        if (cropBox != null)
                                        {
                                            var cropMin = cropBox.Min;
                                            var cropMax = cropBox.Max;
                                            if (intersectionPoint.X < cropMin.X || intersectionPoint.X > cropMax.X ||
                                                intersectionPoint.Y < cropMin.Y || intersectionPoint.Y > cropMax.Y)
                                            {
                                                skippedCount++;
                                                continue;
                                            }
                                        }

                                        string horizontalGridName = horizontalGrid.Name;
                                        string verticalGridName = verticalGrid.Name;

                                        // Skip if an existing label already covers this intersection
                                        string existingLabelKey = $"{horizontalGridName}|{verticalGridName}";
                                        if (skipExisting && existingKeys.Contains(existingLabelKey))
                                        {
                                            skippedCount++;
                                            continue;
                                        }

                                        // Build label text from the template
                                        string labelText;
                                        if (orderHV)
                                            labelText = labelTemplate.Replace("{H}", horizontalGridName).Replace("{V}", verticalGridName);
                                        else
                                            labelText = labelTemplate.Replace("{H}", verticalGridName).Replace("{V}", horizontalGridName);

                                        // Create element at the raw intersection point (alignment applied in Phase 3)
                                        Element placedElement = null;

                                        if (useTextNote)
                                        {
                                            placedElement = TextNote.Create(doc, view.Id, intersectionPoint,
                                                labelText, selectedTextType.Id);
                                        }
                                        else
                                        {
                                            var familyInstance = doc.Create.NewFamilyInstance(
                                                intersectionPoint, selectedSymbol, view);
                                            var textParameter = familyInstance.LookupParameter(selectedParam.Name);
                                            if (textParameter != null && !textParameter.IsReadOnly)
                                                textParameter.Set(labelText);
                                            placedElement = familyInstance;
                                        }

                                        if (placedElement == null) continue;

                                        // Tag with Extensible Storage for later identification
                                        TagElement(placedElement, ViewIdToString(view.Id), horizontalGridName, verticalGridName);

                                        // Queue for post-creation alignment (bbox measurement + rotation + move)
                                        bool shouldRotateThisLabel = rotateDiagonalLabels
                                            && !pairingSnapshot.IsOrthogonalPairing;

                                        pendingLabelAlignments.Add(new PendingLabelAlignment
                                        {
                                            PlacedElement = placedElement,
                                            IntersectionPoint = intersectionPoint,
                                            HorizontalGridCurve = horizontalCurve,
                                            ShouldRotate = shouldRotateThisLabel,
                                            HorizontalGridName = horizontalGridName,
                                            VerticalGridName = verticalGridName
                                        });

                                        placedCount++;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    errorCount++;
                                    errorDetails += $"{horizontalGrid.Name} x {verticalGrid.Name}: {ex.Message}\n";
                                }
                            }
                        }
                    }

                    // ── Phase 2: Single Regenerate to make all bounding boxes available ──
                    // This replaces per-element Regenerate calls, significantly improving
                    // performance for large grid counts (e.g., 100+ intersections).
                    if (pendingLabelAlignments.Count > 0)
                        doc.Regenerate();

                    // ── Phase 3: Align each element using its bounding box ──
                    // Measure bbox, compute alignment offset, optionally rotate for diagonal
                    // grids, then move to final position.
                    foreach (var pendingAlignment in pendingLabelAlignments)
                    {
                        try
                        {
                            var placedElement = pendingAlignment.PlacedElement;
                            var intersectionPoint = pendingAlignment.IntersectionPoint;

                            BoundingBoxXYZ elementBoundingBox = placedElement.get_BoundingBox(view);

                            // Compute how far each edge of the bbox is from the intersection
                            double distanceToLeftEdge = elementBoundingBox != null
                                ? elementBoundingBox.Min.X - intersectionPoint.X : 0;
                            double distanceToRightEdge = elementBoundingBox != null
                                ? elementBoundingBox.Max.X - intersectionPoint.X : 0;
                            double distanceToBottomEdge = elementBoundingBox != null
                                ? elementBoundingBox.Min.Y - intersectionPoint.Y : 0;
                            double distanceToTopEdge = elementBoundingBox != null
                                ? elementBoundingBox.Max.Y - intersectionPoint.Y : 0;

                            // Compute the alignment move vector.
                            // Both modes use bbox edge alignment to position the label on the
                            // correct side of the grid lines. The difference is the padding:
                            // Auto mode uses a view-scale-based padding; Manual mode uses the
                            // user-specified X/Y offsets as the padding distance.
                            double alignmentMoveX, alignmentMoveY;
                            double paddingX = autoOffset ? paddingModelUnits : manualModelOffsetX;
                            double paddingY = autoOffset ? paddingModelUnits : manualModelOffsetY;

                            if (alignRightOfGrid)
                                alignmentMoveX = paddingX - distanceToLeftEdge;
                            else
                                alignmentMoveX = -paddingX - distanceToRightEdge;

                            if (alignAboveGrid)
                                alignmentMoveY = paddingY - distanceToBottomEdge;
                            else
                                alignmentMoveY = -paddingY - distanceToTopEdge;

                            // Rotate for diagonal grids
                            if (pendingAlignment.ShouldRotate)
                            {
                                // Compute rotation angle from the H-token grid's tangent direction
                                XYZ horizontalGridTangent = GetCurveTangentDirection(pendingAlignment.HorizontalGridCurve);
                                double rotationAngleRadians = Math.Atan2(horizontalGridTangent.Y, horizontalGridTangent.X);

                                // Keep text readable (not upside down): constrain to ±90° range
                                if (rotationAngleRadians > Math.PI / 2.0)
                                    rotationAngleRadians -= Math.PI;
                                if (rotationAngleRadians < -Math.PI / 2.0)
                                    rotationAngleRadians += Math.PI;

                                // Rotate the alignment move vector to match the grid direction
                                double cosAngle = Math.Cos(rotationAngleRadians);
                                double sinAngle = Math.Sin(rotationAngleRadians);
                                double rotatedMoveX = alignmentMoveX * cosAngle - alignmentMoveY * sinAngle;
                                double rotatedMoveY = alignmentMoveX * sinAngle + alignmentMoveY * cosAngle;
                                alignmentMoveX = rotatedMoveX;
                                alignmentMoveY = rotatedMoveY;

                                // Rotate the element around the intersection point
                                Line rotationAxis = Line.CreateBound(
                                    intersectionPoint,
                                    intersectionPoint + XYZ.BasisZ);
                                ElementTransformUtils.RotateElement(
                                    doc, placedElement.Id, rotationAxis, rotationAngleRadians);
                            }

                            // Move element to aligned position
                            XYZ alignmentMoveVector = new XYZ(alignmentMoveX, alignmentMoveY, 0);
                            ElementTransformUtils.MoveElement(doc, placedElement.Id, alignmentMoveVector);
                        }
                        catch (Exception ex)
                        {
                            errorCount++;
                            errorDetails += $"{pendingAlignment.HorizontalGridName} x {pendingAlignment.VerticalGridName} (align): {ex.Message}\n";
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

            // Update results on the WPF thread
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

            // Capture all grid names across all displayed groups for scope-aware filtering
            HashSet<string> allDisplayedGridNames = null;
            bool isScopeAllVisible = true;

            Dispatcher.Invoke(() =>
            {
                isScopeAllVisible = RbScopeAllVisible?.IsChecked == true;
                if (!isScopeAllVisible)
                {
                    // Collect grid names from ALL groups (not just H/V) into a single set
                    allDisplayedGridNames = new HashSet<string>();
                    foreach (var gridGroup in _allGridGroups)
                    {
                        foreach (var gridItem in gridGroup.GridItems)
                            allDisplayedGridNames.Add(gridItem.GridName);
                    }
                }
            });

            int deletedCount = 0;

            using (Transaction tx = new Transaction(doc, "GridCoords: Delete Labels"))
            {
                tx.Start();
                try
                {
                    var taggedElements = FindTaggedElements(doc, view);

                    // Filter to only labels whose grid names exist in displayed groups
                    if (!isScopeAllVisible && allDisplayedGridNames != null)
                    {
                        var extensibleStorageSchema = Schema.Lookup(SchemaGuid);
                        taggedElements = taggedElements.Where(element =>
                        {
                            var entity = element.GetEntity(extensibleStorageSchema);
                            if (!entity.IsValid()) return false;
                            string storedHGridName = entity.Get<string>(FieldHGridName);
                            string storedVGridName = entity.Get<string>(FieldVGridName);
                            return allDisplayedGridNames.Contains(storedHGridName)
                                && allDisplayedGridNames.Contains(storedVGridName);
                        }).ToList();
                    }

                    deletedCount = taggedElements.Count;
                    foreach (var element in taggedElements)
                        doc.Delete(element.Id);
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

            int totalDeletedCount = deletedCount;
            Dispatcher.BeginInvoke(new Action(() =>
            {
                ShowResults(0, totalDeletedCount, 0, 0, "");
                SetStatus($"{totalDeletedCount} labels deleted.");
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

        /// <summary>
        /// Holds data for a newly created label element that needs post-creation alignment
        /// (bounding box measurement, optional rotation, and position adjustment).
        /// Used by the batched placement pipeline to separate element creation from alignment.
        /// </summary>
        private class PendingLabelAlignment
        {
            public Element PlacedElement { get; set; }
            public XYZ IntersectionPoint { get; set; }
            public Curve HorizontalGridCurve { get; set; }
            public bool ShouldRotate { get; set; }
            public string HorizontalGridName { get; set; }
            public string VerticalGridName { get; set; }
        }
    }

    // ── Data models ──

    /// <summary>
    /// Represents a single grid row in the checkbox lists. Tracks selection state,
    /// grid name, orientation classification, and angle for grouping.
    /// </summary>
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

        /// <summary>
        /// Classification of this grid: "Horizontal", "Vertical", "Diagonal", or "Curved".
        /// May be reassigned during angle-based clustering.
        /// </summary>
        public string Orientation
        {
            get => _orientation;
            set { _orientation = value; OnPropertyChanged(); }
        }

        public ElementId GridId { get; set; }

        /// <summary>
        /// Angle in degrees (0–180) relative to the view's horizontal axis.
        /// 0° = horizontal, 90° = vertical. Curved grids use -1.
        /// </summary>
        public double AngleFromViewHorizontal { get; set; } = -1;

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
    /// Represents a group of grids that share a similar angle (e.g., all ~0° grids, all ~90° grids, all ~45° grids).
    /// Encapsulates the collection, label, and SelectAll logic for one angle cluster.
    /// </summary>
    public class GridGroup : INotifyPropertyChanged
    {
        private bool _isSelectAllChecked = true;
        private bool _suppressSelectAllEvents;

        /// <summary>
        /// The grids belonging to this angle group.
        /// </summary>
        public ObservableCollection<GridRowItem> GridItems { get; } = new ObservableCollection<GridRowItem>();

        /// <summary>
        /// Display label for this group, e.g., "Horizontal (0°)", "Diagonal (~45°)".
        /// </summary>
        public string GroupDisplayLabel { get; set; }

        /// <summary>
        /// The representative angle for this group in degrees (0–180), used for sorting and token assignment.
        /// -1 for curved-only groups.
        /// </summary>
        public double RepresentativeAngleDegrees { get; set; }

        /// <summary>
        /// Whether this is one of the two standard groups (Horizontal ~0° or Vertical ~90°).
        /// Standard groups use the XAML-defined expanders; non-standard groups are dynamically generated.
        /// </summary>
        public bool IsStandardGroup { get; set; }

        /// <summary>
        /// A short classification tag: "Horizontal", "Vertical", "Diagonal", or "Curved".
        /// Used for token assignment logic and header generation.
        /// </summary>
        public string GroupClassification { get; set; }

        /// <summary>
        /// Tracks the Select All checkbox state for this group with built-in event suppression.
        /// </summary>
        public bool IsSelectAllChecked
        {
            get => _isSelectAllChecked;
            set { _isSelectAllChecked = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// Applies the Select All state to every grid item in this group, suppressing cascading events.
        /// </summary>
        public void ApplySelectAllToItems(bool isChecked)
        {
            _suppressSelectAllEvents = true;
            foreach (var gridItem in GridItems)
                gridItem.IsSelected = isChecked;
            _suppressSelectAllEvents = false;
        }

        /// <summary>
        /// Returns true if event suppression is active (to prevent cascading checkbox events).
        /// </summary>
        public bool IsSuppressingEvents => _suppressSelectAllEvents;

        /// <summary>
        /// Count of selected grids in this group.
        /// </summary>
        public int SelectedGridCount => GridItems.Count(gridItem => gridItem.IsSelected);

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    /// <summary>
    /// Represents a pairing between two grid groups for intersection computation.
    /// For example, Horizontal × Vertical, or Diagonal ~30° × Diagonal ~120°.
    /// </summary>
    public class IntersectionPairing : INotifyPropertyChanged
    {
        private bool _isEnabled = true;

        /// <summary>
        /// The first group in the pairing (assigned the {H} token — the "more horizontal" group).
        /// </summary>
        public GridGroup GroupForHToken { get; set; }

        /// <summary>
        /// The second group in the pairing (assigned the {V} token — the "more vertical" group).
        /// </summary>
        public GridGroup GroupForVToken { get; set; }

        /// <summary>
        /// Whether this pairing is enabled for placement. Users can disable specific pairings.
        /// </summary>
        public bool IsEnabled
        {
            get => _isEnabled;
            set { _isEnabled = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// Display label for this pairing, e.g., "Horizontal × Vertical" or "~30° × ~120°".
        /// </summary>
        public string PairingDisplayLabel { get; set; }

        /// <summary>
        /// Estimated intersection count: selected grids in H group × selected grids in V group.
        /// </summary>
        public int EstimatedIntersectionCount =>
            GroupForHToken.SelectedGridCount * GroupForVToken.SelectedGridCount;

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    /// <summary>
    /// Lightweight snapshot of a pairing's selected grids and group angles,
    /// captured on the UI thread for safe use on the Revit external event thread.
    /// </summary>
    public class PairingSnapshot
    {
        public List<GridRowItem> HorizontalGridItems { get; set; }
        public List<GridRowItem> VerticalGridItems { get; set; }

        /// <summary>
        /// Representative angle in degrees of the H-token group (the "more horizontal" group).
        /// 0 for standard Horizontal, 90 for standard Vertical, other values for diagonal.
        /// </summary>
        public double HGroupAngleDegrees { get; set; }

        /// <summary>
        /// Representative angle in degrees of the V-token group (the "more vertical" group).
        /// </summary>
        public double VGroupAngleDegrees { get; set; }

        /// <summary>
        /// True if this pairing is between the standard Horizontal (0°) and Vertical (90°) groups.
        /// Used for optimized orthogonal offset placement.
        /// </summary>
        public bool IsOrthogonalPairing =>
            Math.Abs(HGroupAngleDegrees) < 10 && Math.Abs(VGroupAngleDegrees - 90) < 10;
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
