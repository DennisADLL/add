using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;

namespace add
{
    [Transaction(TransactionMode.Manual)]
    public class AddDivideWallApiladoCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                IList<Reference> refs = uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    new StackedWallSelectionFilter(),
                    "Selecciona muros apilados");

                List<Wall> walls = new List<Wall>();

                foreach (Reference r in refs)
                {
                    Wall wall = doc.GetElement(r) as Wall;

                    if (wall != null && wall.IsStackedWall)
                    {
                        walls.Add(wall);
                    }

                    using (Transaction t = new Transaction(doc, "Convert Stacked Wall to Normal Walls"))
                    {
                        t.Start();

                        IList<ElementId> memberIds = wall.GetStackedWallMemberIds();

                        if (memberIds == null || memberIds.Count == 0)
                        {
                            TaskDialog.Show("Error", "The stacked wall has no member walls.");
                            t.RollBack();
                            return Result.Failed;
                        }

                        // Copy member walls in the same location
                        ICollection<ElementId> newWallIds =
                            ElementTransformUtils.CopyElements(doc, memberIds, XYZ.Zero);

                        // Delete original stacked wall
                        doc.Delete(wall.Id);

                        t.Commit();

                        //TaskDialog.Show("Success",
                        //    $"Stacked wall converted.\nNew walls created: {newWallIds.Count}");
                    }
                }                              
                                

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (System.Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
        public class StackedWallSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                Wall wall = elem as Wall;
                return wall != null && wall.IsStackedWall;
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }
    }
}
