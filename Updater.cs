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

            using (var transaction = new Transaction(doc))
            {
                if (transaction.Start("Replace null with zero") == TransactionStatus.Started)
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

                transaction.Commit();


                // Above now auto commits. 
                //// Ask the end user whether the changes are to be committed or not
                //var taskDialog = new TaskDialog("Revit");
                //taskDialog.MainContent = "Click either [OK] to Commit, or [Cancel] to Roll back the transaction.";
                //var buttons = TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel;
                //taskDialog.CommonButtons = buttons;

                //if (TaskDialogResult.Ok == taskDialog.Show())
                //{
                //    // For many various reasons, a transaction may not be committed
                //    // if the changes made during the transaction do not result a valid model.
                //    // If committing a transaction fails or is canceled by the end user,
                //    // the resulting status would be RolledBack instead of Committed.
                //    if (TransactionStatus.Committed != transaction.Commit())
                //        TaskDialog.Show("Failure", "Transaction could not be committed");
                //}
                //else
                //{
                //    transaction.RollBack();
                //}
            }

            return Result.Succeeded;
        }
    }
}