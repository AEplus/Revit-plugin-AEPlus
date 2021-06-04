using System.Collections.Generic;
using System.Linq;
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

            using (var trans = new Transaction(doc))
            {
                if (trans.Start("Replace null with zero") == TransactionStatus.Started)
                {
                    BuiltInCategory[] bics =
                    {
                        BuiltInCategory.OST_CableTray,
                        BuiltInCategory.OST_CableTrayFitting,
                        BuiltInCategory.OST_Conduit,
                        BuiltInCategory.OST_ConduitFitting,
                        BuiltInCategory.OST_DuctCurves,
                        BuiltInCategory.OST_DuctFitting,
                        BuiltInCategory.OST_DuctTerminal,
                        BuiltInCategory.OST_ElectricalEquipment,
                        BuiltInCategory.OST_ElectricalFixtures,
                        BuiltInCategory.OST_LightingDevices,
                        BuiltInCategory.OST_LightingFixtures,
                        BuiltInCategory.OST_MechanicalEquipment,
                        BuiltInCategory.OST_PipeCurves,
                        BuiltInCategory.OST_PipeFitting,
                        BuiltInCategory.OST_PlumbingFixtures,
                        BuiltInCategory.OST_SpecialityEquipment,
                        BuiltInCategory.OST_Sprinklers,
                        BuiltInCategory.OST_Wire
                    };

                    IList<ElementFilter> a
                        = new List<ElementFilter>(bics.Count());

                    foreach (var bic in bics) a.Add(new ElementCategoryFilter(bic));
                    var categoryFilter
                        = new LogicalOrFilter(a);

                    ElementFilter f = new LogicalAndFilter(
                        categoryFilter,
                        new ElementIsElementTypeFilter(true));

                    var collector
                        = new FilteredElementCollector(doc)
                            .WherePasses(f);

                    ElementId id = null;

                    foreach (var e in collector)
                    {
                        var parameter = e.LookupParameter("AE Opmeting");
                        if (parameter.AsDouble() == 0) parameter.Set(0);

                        // var P = doc.GetElement(e.Id).LookupParameter("AE Opmeting");
                        // if (P != null || P.AsValueString() != "0" || P.AsValueString() != null) P.Set("0");
                    }
                }

                trans.Commit();
            }

            return Result.Succeeded;
        }
    }
}