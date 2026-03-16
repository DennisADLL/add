using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace add
{
    [Transaction(TransactionMode.Manual)]
    public class AddCutWallLinkPipeCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // 1) Seleccionar varios tubos de vínculos
                IList<Reference> pipeRefs = uidoc.Selection.PickObjects(
                    ObjectType.LinkedElement,
                    new LinkedPipeSelectionFilter(),
                    "Selecciona tubos de un vínculo");

                if (pipeRefs == null || pipeRefs.Count == 0)
                {
                    message = "No se seleccionaron tubos del vínculo.";
                    return Result.Cancelled;
                }

                // 2) Seleccionar varios muros del host
                IList<Reference> wallRefs = uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    new HostWallSelectionFilter(),
                    "Selecciona muros del modelo anfitrión");

                if (wallRefs == null || wallRefs.Count == 0)
                {
                    message = "No se seleccionaron muros.";
                    return Result.Cancelled;
                }

                // Agrupar tubos seleccionados con su link instance y transform
                List<LinkedPipeData> linkedPipes = new List<LinkedPipeData>();

                foreach (Reference pipeRef in pipeRefs)
                {
                    RevitLinkInstance linkInstance = doc.GetElement(pipeRef.ElementId) as RevitLinkInstance;
                    if (linkInstance == null)
                        continue;

                    Document linkDoc = linkInstance.GetLinkDocument();
                    if (linkDoc == null)
                        continue;

                    Element linkedElement = linkDoc.GetElement(pipeRef.LinkedElementId);
                    Pipe pipe = linkedElement as Pipe;
                    if (pipe == null)
                        continue;

                    linkedPipes.Add(new LinkedPipeData
                    {
                        LinkInstance = linkInstance,
                        LinkDocument = linkDoc,
                        Pipe = pipe,
                        TransformToHost = linkInstance.GetTotalTransform()
                    });
                }

                if (linkedPipes.Count == 0)
                {
                    message = "No se obtuvo ninguna tubería válida del vínculo.";
                    return Result.Failed;
                }

                List<Wall> walls = new List<Wall>();

                foreach (Reference wallRef in wallRefs)
                {
                    Wall wall = doc.GetElement(wallRef) as Wall;
                    if (wall != null)
                        walls.Add(wall);
                }

                if (walls.Count == 0)
                {
                    message = "No se obtuvo ningún muro válido.";
                    return Result.Failed;
                }

                int intersectionsFound = 0;
                int editableWalls = 0;

                

                    foreach (Wall wall in walls)
                    {
                        // Solo muros rectos compatibles
                        if (!(wall.Location is LocationCurve wallLoc) || !(wallLoc.Curve is Line wallLine))
                            continue;

                        Solid wallSolid = GetMainSolid(wall, Transform.Identity);
                        if (wallSolid == null || wallSolid.Volume <= 1e-9)
                            continue;

                        foreach (LinkedPipeData pipeData in linkedPipes)
                        {
                            Pipe pipe = pipeData.Pipe;
                            Transform linkTransform = pipeData.TransformToHost;

                            Solid pipeSolidInHost = GetMainSolid(pipe, linkTransform);
                            if (pipeSolidInHost == null || pipeSolidInHost.Volume <= 1e-9)
                                continue;

                            Solid intersectionSolid = null;

                            try
                            {
                                intersectionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(
                                    wallSolid,
                                    pipeSolidInHost,
                                    BooleanOperationsType.Intersect);
                            }
                            catch
                            {
                                continue;
                            }

                            if (intersectionSolid == null || intersectionSolid.Volume <= 1e-9)
                                continue;

                            intersectionsFound++;

                        XYZ hitPoint = GetPipeAxisWallPlaneIntersection(wall, pipe, linkTransform);
                        if (hitPoint == null)
                            continue;

                        // 1) Verificar si el muro admite ProfileSketch
                        if (!wall.CanHaveProfileSketch())
                                continue;

                            // 2) Crear ProfileSketch si aún no existe
                            if (wall.SketchId == ElementId.InvalidElementId)
                            {
                                using (Transaction txCreate = new Transaction(doc, "Crear ProfileSketch"))
                                {
                                    txCreate.Start();
                                    wall.CreateProfileSketch();
                                    txCreate.Commit();
                                }
                            }

                            if (wall.SketchId == ElementId.InvalidElementId)
                                continue;

                            // 3) Entrar a SketchEditScope
                            SketchEditScope scope = new SketchEditScope(doc, "Editar perfil muro");
                            scope.Start(wall.SketchId);

                            using (Transaction txEdit = new Transaction(doc, "Modificar perfil"))
                            {
                                txEdit.Start();

                                Sketch sketch = doc.GetElement(wall.SketchId) as Sketch;
                                if (sketch != null)
                                {
                                    SketchPlane sp = sketch.SketchPlane;
                                    Plane plane = sp.GetPlane();

                                    double radius = UnitUtils.ConvertToInternalUnits(0.15, UnitTypeId.Meters);

                                    // proyectar el centro al plano del sketch
                                    XYZ centerOnPlane = ProjectPointToPlane(hitPoint, plane);

                                    Arc circle = Arc.Create(
                                        centerOnPlane,
                                        radius,
                                        0,
                                        2 * Math.PI,
                                        plane.XVec,
                                        plane.YVec);

                                     doc.Create.NewModelCurve(circle, sp);

                                     editableWalls++;
                                }

                                txEdit.Commit();
                            }

                            scope.Commit(new SimpleFailuresPreprocessor());



                        }
                    }

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private static double GetPipeOuterDiameter(Pipe pipe)
        {
            Parameter p = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
            if (p != null && p.HasValue)
                return p.AsDouble();

            p = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
            if (p != null && p.HasValue)
                return p.AsDouble();

            return 0.0;
        }

        private static Solid GetMainSolid(Element element, Transform extraTransform)
        {
            Options opt = new Options
            {
                ComputeReferences = false,
                DetailLevel = ViewDetailLevel.Fine,
                IncludeNonVisibleObjects = false
            };

            GeometryElement geo = element.get_Geometry(opt);
            if (geo == null) return null;

            List<Solid> solids = new List<Solid>();

            foreach (GeometryObject obj in geo)
            {
                ExtractSolids(obj, extraTransform, solids);
            }

            return solids
                .Where(s => s != null && s.Volume > 1e-9)
                .OrderByDescending(s => s.Volume)
                .FirstOrDefault();
        }

        private static void ExtractSolids(GeometryObject obj, Transform currentTransform, List<Solid> solids)
        {
            if (obj is Solid solid)
            {
                if (solid.Volume > 1e-9)
                {
                    solids.Add(SolidUtils.CreateTransformed(solid, currentTransform));
                }
                return;
            }

            if (obj is GeometryInstance gi)
            {
                Transform nested = currentTransform.Multiply(gi.Transform);
                GeometryElement instGeo = gi.GetInstanceGeometry();

                foreach (GeometryObject nestedObj in instGeo)
                {
                    ExtractSolids(nestedObj, nested, solids);
                }
            }
        }

        private static XYZ GetSolidCentroidApprox(Solid solid)
        {
            BoundingBoxXYZ bb = solid.GetBoundingBox();
            if (bb == null) return null;

            return (bb.Min + bb.Max) * 0.5;
        }

        private static bool OpeningAlreadyExistsNear(Document doc, Wall wall, XYZ center, double width, double height)
        {
            ICollection<ElementId> inserts = wall.FindInserts(true, true, true, true);
            if (inserts == null || inserts.Count == 0)
                return false;

            double tol = Math.Max(width, height) * 0.5;

            foreach (ElementId id in inserts)
            {
                Element e = doc.GetElement(id);
                if (e is Opening opening)
                {
                    BoundingBoxXYZ bb = opening.get_BoundingBox(null);
                    if (bb == null) continue;

                    XYZ c = (bb.Min + bb.Max) * 0.5;
                    if (c.DistanceTo(center) < tol)
                        return true;
                }
            }

            return false;
        }

        private static XYZ ProjectPointToPlane(XYZ point, Plane plane)
        {
            XYZ v = point - plane.Origin;
            double distance = v.DotProduct(plane.Normal);
            return point - distance * plane.Normal;
        }

        private class LinkedPipeData
        {
            public RevitLinkInstance LinkInstance { get; set; }
            public Document LinkDocument { get; set; }
            public Pipe Pipe { get; set; }
            public Transform TransformToHost { get; set; }
        }

        private static XYZ GetPipeAxisWallPlaneIntersection(Wall wall, Pipe pipe, Transform linkTransform)
        {
            LocationCurve wallLoc = wall.Location as LocationCurve;
            if (wallLoc == null) return null;

            Line wallLine = wallLoc.Curve as Line;
            if (wallLine == null) return null;

            LocationCurve pipeLoc = pipe.Location as LocationCurve;
            if (pipeLoc == null) return null;

            Curve pipeCurve = pipeLoc.Curve;
            if (pipeCurve == null) return null;

            XYZ pipeP0 = linkTransform.OfPoint(pipeCurve.GetEndPoint(0));
            XYZ pipeP1 = linkTransform.OfPoint(pipeCurve.GetEndPoint(1));
            XYZ pipeDir = (pipeP1 - pipeP0).Normalize();

            XYZ wallDir = wallLine.Direction.Normalize();
            XYZ up = XYZ.BasisZ;

            XYZ wallNormal = wallDir.CrossProduct(up);
            if (wallNormal.GetLength() < 1e-9)
                return null;

            wallNormal = wallNormal.Normalize();

            // plano medio del muro
            XYZ wallOrigin = wallLine.GetEndPoint(0);
            Plane wallPlane = Plane.CreateByNormalAndOrigin(wallNormal, wallOrigin);

            double denom = pipeDir.DotProduct(wallPlane.Normal);
            if (Math.Abs(denom) < 1e-9)
                return null;

            double t = (wallPlane.Origin - pipeP0).DotProduct(wallPlane.Normal) / denom;

            XYZ hitPoint = pipeP0 + t * pipeDir;
            return hitPoint;
        }
    }

    public class LinkedPipeSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is RevitLinkInstance;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return true;
        }
    }

    public class HostWallSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is Wall;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
    public class SimpleFailuresPreprocessor : IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
        {
            IList<FailureMessageAccessor> failureMessages = failuresAccessor.GetFailureMessages();

            foreach (FailureMessageAccessor fma in failureMessages)
            {
                FailureSeverity severity = fma.GetSeverity();

                if (severity == FailureSeverity.Warning)
                {
                    failuresAccessor.DeleteWarning(fma);
                }
            }

            return FailureProcessingResult.Continue;
        }
    }
}