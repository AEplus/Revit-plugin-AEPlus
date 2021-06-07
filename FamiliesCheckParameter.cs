using System.IO;
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
            var uiapp = commandData.Application;
            var uidoc = uiapp.ActiveUIDocument;
            var app = uiapp.Application;
            var doc = uidoc.Document;

            var assemblyPath = Path.GetDirectoryName("C:/Documents");

            var filename = "C:/Documents";


            // bool LoadFamily(string filename, IFamilyLoadOptions familyLoadOptions, out Family family);

            //if (doc.IsFamilyDocument)
            //{
            //    var f = doc.OwnerFamily;
            //    var manager = doc.FamilyManager;
            //    //Check if parameters have already been added to Family and it's Symbols.
            //    //If not then adding parameters using Transaction.
            //    if (manager.get_Parameter("DMS_DOCNUMBER") == null)
            //        using (var trans = new Transaction(doc, "PARAM_SET"))
            //        {
            //            trans.Start();
            //            manager.AddParameter("AE GeoITLink", BuiltInParameterGroup.PG_GENERAL, ParameterType.Integer,
            //                true);
            //            trans.Commit();
            //        }

            //    //Create FamilyParameter reference objects after parameters have been added or checked if added  
            //    var AEGeoITLink = manager.get_Parameter("AE GeoITLink");

            //    //Setting parameter values for every Symbol in the family
            //    foreach (FamilyType ft in manager.Types)
            //        if (!MaterialHasDocument(ft, manager))
            //        {
            //            DocumentObj docObj = CreateDMSDocument("RVT");
            //            using (var trans = new Transaction(doc, "SET_PARAM"))
            //            {
            //                trans.Start();
            //                manager.CurrentType = ft;
            //                manager.Set(AEGeoITLink, docObj.docNumber);
            //                trans.Commit();
            //            }


            return Result.Succeeded;
        }
    }
}