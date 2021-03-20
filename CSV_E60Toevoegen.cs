using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Globalization;
using System.IO;

namespace MyRevitCommands
{
    [TransactionAttribute(TransactionMode.ReadOnly)]
    public class CSV_E60Toevoegen : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            ArrayList revitSchedules = new ArrayList();
            revitSchedules.Add("AE_M50");
            revitSchedules.Add("AE_M52");
            revitSchedules.Add("AE_M57_ Ventilatieroosters");
            revitSchedules.Add("AE_M57_Toestellen VENT");
            revitSchedules.Add("AE_M50_Toestellen HVAC coll");

            return new GenericToevoegen().GenericExecute(commandData, ref message, elements, @"c:\\temp\\E_60TEST\", revitSchedules);
        }
    }
}