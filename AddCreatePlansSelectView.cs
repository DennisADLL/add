using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace add
{
    [Transaction(TransactionMode.Manual)]
    public class AddCreatePlansSelectView : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            List<ElementId> elementsIdSelection = uidoc.Selection.GetElementIds().ToList();                     

            //List<View> elementViews = elementsIdSelection.Select(r => doc.GetElement(r)).OfType<View>().ToList();

            List<ViewSheet> elementViewSheet = elementsIdSelection.Select(r => doc.GetElement(r)).OfType<ViewSheet>().ToList();

            TaskDialog.Show("Views", elementViewSheet.Count.ToString());

            foreach (ViewSheet viewSheet in elementViewSheet)
            {
                using (Transaction tran = new Transaction(doc, "Create plans of views"))
                {
                    tran.Start();

                    ElementId duplicatedViewSheetId = viewSheet.Duplicate(SheetDuplicateOption.DuplicateSheetWithViewsOnly);

                    ViewSheet elementoViewSheetDuplicated = doc.GetElement(duplicatedViewSheetId) as ViewSheet;

                    tran.Commit();
                }
            }
                       

            return Result.Succeeded;
        }

        public static double MnToFt(int mm)
        {
            return UnitUtils.ConvertToInternalUnits(mm, UnitTypeId.Millimeters);
        }
    }

    
}
