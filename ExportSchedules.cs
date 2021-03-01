using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyRevitCommands
{
    [TransactionAttribute(TransactionMode.ReadOnly)]
    public class ExportingSchedules : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            //Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;

            //Get Document
            Document doc = uidoc.Document;

            FilteredElementCollector col = new FilteredElementCollector(doc)
            .OfClass(typeof(ViewSchedule));

            ViewScheduleExportOptions opt = new ViewScheduleExportOptions();

            foreach (ViewSchedule vs in col)
            {
                Directory.CreateDirectory(@"c:\\temp\\test");
                vs.Export(@"c:\\temp\\test", vs.Name + ".csv", opt);
            }

            return Result.Succeeded;
        }
    }
}
