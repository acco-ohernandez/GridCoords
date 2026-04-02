using GridCoords.Common;
using GridCoords.Forms;

namespace GridCoords
{
    [Transaction(TransactionMode.Manual)]
    public class Cmd_GridCoords : IExternalCommand
    {
        private static GridCoords_Form _form;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;

            if (uidoc == null)
            {
                TaskDialog.Show("Grid Coords", "No active document.");
                return Result.Cancelled;
            }

            // If form already open, reactivate it
            if (_form != null && _form.IsVisible)
            {
                Utils.PositionWindowCenterRight(_form, uiapp.MainWindowHandle);
                _form.Activate();
                return Result.Succeeded;
            }

            Document doc = uidoc.Document;
            View activeView = uidoc.ActiveView;

            // Validate view type
            if (!(activeView is ViewPlan))
            {
                TaskDialog.Show("Grid Coords", "Please open a plan view (Floor Plan, Ceiling Plan, or Area Plan).");
                return Result.Cancelled;
            }

            // Collect grid data on the Revit thread
            var gridItems = GridCoords_Form.CollectGridData(doc, activeView);
            string viewName = activeView.Name;
            ElementId viewId = activeView.Id;

            // Create handler + external event
            var handler = new GridCoords_EventHandler();
            ExternalEvent externalEvent = ExternalEvent.Create(handler);

            // Create and show modeless form
            _form = new GridCoords_Form(uiapp, externalEvent, handler, gridItems, viewName, viewId);
            _form.Show();

            return Result.Succeeded;
        }

        internal static PushButtonData GetButtonData()
        {
            string buttonInternalName = "btnCmd_GridCoords";
            string buttonTitle = "Grid Coords";

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "Place coordinate labels at grid intersections.");

            return myButtonData.Data;
        }
    }
}
