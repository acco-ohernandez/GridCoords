namespace GridCoords.Common
{
    public class GridCoords_EventHandler : IExternalEventHandler
    {
        public Action<UIApplication> HandlerAction { get; set; }

        public void Execute(UIApplication app)
        {
            try
            {
                HandlerAction?.Invoke(app);
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // User cancelled — silently ignore.
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"GridCoords handler error: {ex}");
                TaskDialog.Show("Grid Coords – Error", ex.Message);
            }
            finally
            {
                HandlerAction = null;
            }
        }

        public string GetName() => nameof(GridCoords_EventHandler);
    }
}
