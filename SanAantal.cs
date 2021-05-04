using System;
using System.Collections;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace MyRevitCommands
{
    [TransactionAttribute(TransactionMode.ReadOnly)]
    public class SanAantal : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var revitSchedules = new ArrayList();
                // revitSchedules.Add("AE_M56");
                // revitSchedules.Add("ventilatierooster");
                revitSchedules.Add("SAN A");

                var fileName = GetType().Name;

                return new GenericToevoegen().GenericExecute(commandData, ref message, elements, @"c:\\temp\\Sanitair\",
                    revitSchedules, fileName);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }
    }
}