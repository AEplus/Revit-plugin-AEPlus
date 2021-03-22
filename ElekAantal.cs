using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections;
namespace MyRevitCommands
{
    [TransactionAttribute(TransactionMode.ReadOnly)]
    public class ElekAantal : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            ArrayList revitSchedules = new ArrayList();
            revitSchedules.Add("AE_E60");

            string fileName = this.GetType().Name;

            return new GenericToevoegen().GenericExecute(commandData, ref message, elements, @"c:\\temp\\Elektriciteit\", revitSchedules, fileName);
        }
    }
}