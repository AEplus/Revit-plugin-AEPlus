using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace MyRevitCommands
{
    [Transaction(TransactionMode.Manual)]
    public class ImportParameter : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {


            return Result.Succeeded;
        }
    }
}