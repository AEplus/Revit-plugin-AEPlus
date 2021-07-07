using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace MyRevitCommands
{
    [Transaction(TransactionMode.Manual)]
    public class FamiliesCheckParameter : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiapp = commandData.Application;
                var uidoc = uiapp.ActiveUIDocument;
                var app = uiapp.Application;
                var doc = uidoc.Document;

                const string fileName = "Z:/O - Dossiers/O - Bim/stage sebastiaan/Stabiplan Shared Parameter File.txt";
                const string groupname = "00. General";
                const string defname = "AE GeoIT Link";

                uidoc.Application.Application.SharedParametersFilename = fileName;


                var FamilyInstanceFilter = new ElementClassFilter(typeof(FamilyInstance));
                var FamilyInstanceCollector = new FilteredElementCollector(doc);
                ICollection<Element> AllFamilyInstances =
                    FamilyInstanceCollector.WherePasses(FamilyInstanceFilter).ToElements();

                var ParamLst = new List<string>();

                FamilySymbol FmlySmbl;
                Family Fmly;


                var rfaFiles = Directory.GetFiles("C:/Users/sebas/OneDrive/Documenten/2019", "*.rfa",
                    SearchOption.AllDirectories).ToList();
                var i = 0;

                foreach (FamilyInstance FmlyInst in AllFamilyInstances)
                {
                    app.OpenDocumentFile(FmlyInst.Name);
                    if (FmlyInst.LookupParameter("AE GeoIT Link") == null) Fml
                }


                //foreach (FamilyInstance FmlyInst in AllFamilyInstances)
                //    using (var y = new Transaction(doc, "Add shared parameter"))
                //    {
                //        y.Start();
                //        FmlySmbl = FmlyInst.Symbol;
                //        Fmly = FmlySmbl.Family;
                //        var famdoc = doc.EditFamily(Fmly);


                //        // Add Instance Parameter names to list
                //        foreach (Parameter Param in FmlyInst.Parameters) ParamLst.Add(Param.Definition.Name);
                //        if (ParamLst.Contains("AE GeoIT Link") == false)

                //            famdoc.FamilyManager.AddParameter(defname, BuiltInParameterGroup.PG_DATA,
                //                ParameterType.Integer,
                //                true);


                //        // Add Type Parameter names to list
                //        foreach (Parameter Param in FmlySmbl.Parameters) ParamLst.Add(Param.Definition.Name);
                //        if (ParamLst.Contains("AE GeoIT Link") == false)

                //            famdoc.FamilyManager.AddParameter(defname, BuiltInParameterGroup.PG_DATA,
                //                ParameterType.Integer,
                //                true);


                //        // Add Family Type Parameter names to list
                //        foreach (Parameter Param in Fmly.Parameters) ParamLst.Add(Param.Definition.Name);
                //        if (ParamLst.Contains("AE GeoIT Link") == false) ;
                //        famdoc.FamilyManager.AddParameter(defname, BuiltInParameterGroup.PG_DATA,
                //            ParameterType.Integer,
                //            true);
                //        y.Commit();
                //    }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                throw;
            }


            return Result.Succeeded;
        }
    }
}