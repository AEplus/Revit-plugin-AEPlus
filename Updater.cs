using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace MyRevitCommands
{
    [Transaction(TransactionMode.Manual)]
    public class Updater : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var app = commandData.Application;
            var doc = app.ActiveUIDocument.Document;

            var sortedElements
                = new Dictionary<string, List<Element>>();

            // Iterate over all elements, both symbols and 
            // model elements, and them in the dictionary.

            ElementFilter f = new LogicalOrFilter(
                new ElementIsElementTypeFilter(false),
                new ElementIsElementTypeFilter(true));

            var collector
                = new FilteredElementCollector(doc)
                    .WherePasses(f);

            string name;

            foreach (var e in collector)
            {
                var parameter = e.LookupParameter("AE Opmeting");
                // var Opmeting = parameter.AsInteger();

                if (parameter == null) parameter.Set(0);
            }

            return Result.Succeeded;
        }
    }
}