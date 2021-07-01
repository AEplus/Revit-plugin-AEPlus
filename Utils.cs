using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace MyRevitCommands
{
    internal static class Utils
    {
        /// <summary>
        ///     Helper to get all instances for a given category,
        ///     identified either by a built-in category or by a category name.
        /// </summary>
        public static List<Element> GetTargetInstances(
            Document doc,
            object targetCategory)
        {
            List<Element> elements;

            var isName = targetCategory.GetType().Equals(typeof(string));

            if (isName)
            {
                var cat = doc.Settings.Categories.get_Item(targetCategory as string);
                var collector = new FilteredElementCollector(doc);
                collector.OfCategoryId(cat.Id);
                elements = new List<Element>(collector);
            }
            else
            {
                var collector
                    = new FilteredElementCollector(doc)
                        .WhereElementIsNotElementType();

                collector.OfCategory((BuiltInCategory) targetCategory);

                // I removed this to test attaching a shared 
                // parameter to Material elements:
                //
                //var model_elements = from e in collector
                //                     where ( null != e.Category && e.Category.HasMaterialQuantities )
                //                     select e;

                //elements = model_elements.ToList<Element>();

                elements = collector.ToList();
            }

            return elements;
        }
    }
}