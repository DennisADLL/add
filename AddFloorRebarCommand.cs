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
using System.Text;
using System.Threading.Tasks;

namespace add
{
    [Transaction(TransactionMode.Manual)]
    public class AddFloorRebarCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                IList<Reference> refs = uidoc.Selection.PickObjects(ObjectType.Element, new FloorSelectionFilter(), "Selecciona una losas estructurales");
                
                //var col = doc.GetElement(r) as FamilyInstance;

                List<Floor> losas = refs
                    .Select(r => doc.GetElement(r))
                    .OfType<Floor>()
                    .Where(c => c.Category != null &&
                                c.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Floors)
                    .ToList();

                if (losas.Count == 0)
                {
                    TaskDialog.Show("Error", "Los elementos seleccionados no son columnas estructurales.");
                    return Result.Failed;
                }

                // The user can adjust these parameters as needed, or you can implement a UI to input them
                //==========================================                
                double maxSpacingPositiv = MnToFt(150); // 150 mm
                
                double coverConcrete = MnToFt(20); // 40 mm
                //=========================================             
                                
                RebarBarType tieBarType = new FilteredElementCollector(doc)
                    .OfClass(typeof(RebarBarType))
                    .Cast<RebarBarType>()
                    .FirstOrDefault();

                RebarHookType hookType = new FilteredElementCollector(doc)
                    .OfClass(typeof(RebarHookType))
                    .Cast<RebarHookType>()
                    .FirstOrDefault();

                int ok = 0;
                using (Transaction tran = new Transaction(doc, "Crear estribos en columna"))
                {
                    tran.Start();
                    foreach (Floor losa in losas)
                    {
                        FloorRebarService.CreateRebarFloor(
                        doc,
                        losa,
                        tieBarType,
                        hookType,
                        maxSpacingPositiv,
                        coverConcrete,
                        true
                        );

                        FloorRebarService.CreateRebarFloor(
                        doc,
                        losa,
                        tieBarType,
                        hookType,
                        maxSpacingPositiv,
                        coverConcrete,
                        false
                        );
                    }
                    ok++;
                    tran.Commit();
                }

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
                return Result.Failed;
            }

        }

        public static double MnToFt(int mm)
        {
            return UnitUtils.ConvertToInternalUnits(mm, UnitTypeId.Millimeters);
        }
    }

    public class FloorSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem.Category != null && elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Floors;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
    public static class FloorRebarService
    {
        public static void CreateRebarFloor(
            Document doc,
            Floor losa,
            RebarBarType BarType,
            RebarHookType hookType,
            double maxSpacing,
            double concreteCover,
            bool barsAlongX
        )
        {
            if (losa == null) throw new ArgumentNullException(nameof(losa));
            if (BarType == null) throw new ArgumentNullException(nameof(BarType));

            Face caraInferior = losa.ObtenerCaraZ_();
            if (!(caraInferior is PlanarFace pf))
                throw new Exception("La cara inferior de la losa no es planar.");

            XYZ faceNormal = pf.FaceNormal.Normalize();

            IList<CurveLoop> loops = caraInferior.GetEdgesAsCurveLoops();
            if (loops == null || loops.Count == 0)
                throw new Exception("No se encontró contorno en la cara inferior.");

            CurveLoop outerLoop = loops
                .OrderByDescending(cl => Math.Abs(GetCurveLoopPerimeter(cl)))
                .First();

            List<XYZ> pts = outerLoop
                .SelectMany(c => new List<XYZ> { c.GetEndPoint(0), c.GetEndPoint(1) })
                .ToList();

            double minX = pts.Min(p => p.X);
            double maxX = pts.Max(p => p.X);
            double minY = pts.Min(p => p.Y);
            double maxY = pts.Max(p => p.Y);
            double z = pts.Average(p => p.Z);

            // usar el recubrimiento recibido
            double cover = concreteCover;

            double barDia = BarType.get_Parameter(BuiltInParameter.REBAR_BAR_DIAMETER).AsDouble();
            double off = cover + 0.5 * barDia;

            // Ajuste en planta
            minX += cover;
            maxX -= cover;
            minY += cover;
            maxY -= cover;

            if (maxX <= minX || maxY <= minY)
                throw new Exception("El recubrimiento es demasiado grande para la geometría de la losa.");

            // Subir la barra hacia dentro de la losa
            XYZ offsetVec = -faceNormal * off;

            List<Curve> curves = new List<Curve>();
            XYZ normal;

            if (barsAlongX)
            {
                XYZ p1 = new XYZ(minX, minY, z) + offsetVec;
                XYZ p2 = new XYZ(maxX, minY, z) + offsetVec;

                curves.Add(Line.CreateBound(p1, p2));

                // plano vertical XZ
                normal = XYZ.BasisY;
            }
            else
            {
                XYZ p1 = new XYZ(minX, minY, z) + offsetVec;
                XYZ p2 = new XYZ(minX, maxY, z) + offsetVec;

                curves.Add(Line.CreateBound(p1, p2));

                // plano vertical YZ
                normal = XYZ.BasisX;
            }

            Rebar rebar = Rebar.CreateFromCurves(
                doc,
                RebarStyle.Standard,
                BarType,
                hookType,
                hookType,
                losa,
                normal,
                curves,
                RebarHookOrientation.Left,
                RebarHookOrientation.Right,
                true,
                true);

            if (rebar == null)
                throw new Exception("No se pudo crear la armadura longitudinal inferior.");

            double distributionLength = barsAlongX
                ? (maxY - minY)
                : (maxX - minX);

            int numberOfBars = (int)Math.Floor(distributionLength / maxSpacing) + 1;
            if (numberOfBars < 2)
                numberOfBars = 2;

            RebarShapeDrivenAccessor accessor = rebar.GetShapeDrivenAccessor();
            accessor.SetLayoutAsNumberWithSpacing(
                numberOfBars,
                maxSpacing,
                true,
                true,
                true);



        }
        private static double GetCurveLoopPerimeter(CurveLoop loop)
        {
            double length = 0.0;
            foreach (Curve c in loop)
                length += c.Length;
            return length;
        }
        public static Face ObtenerCaraZ_(this Element element)
        {
            Face Cara_ = null;
            Options opciones = new Options();
            GeometryElement geoElemento = element.get_Geometry(opciones);

            foreach (GeometryObject geometryObject in geoElemento)
            {
                if (geometryObject is Solid)
                {
                    Solid solido = geometryObject as Solid;
                    if (solido.Volume > 0)
                    {
                        //listaSolidos.Add(solid);

                        foreach (Face cara in solido.Faces)
                        {
                            if (cara is PlanarFace pf)
                            {
                                XYZ normal = pf.FaceNormal;
                                if (normal.Z < 0)
                                {

                                    Cara_ = cara;
                                }
                            }
                        }
                    }

                }
                else if (geometryObject is GeometryInstance)
                {
                    GeometryInstance geoinstance = geometryObject as GeometryInstance;
                    GeometryElement geoInstanceObject = geoinstance.GetInstanceGeometry();

                    foreach (GeometryObject geomObject in geoInstanceObject)
                    {
                        if (geomObject is Solid)
                        {
                            Solid solido = geomObject as Solid;
                            if (solido.Volume > 0)
                            {
                                //listaSolidos.Add(solido);
                                foreach (Face cara in solido.Faces)
                                {
                                    if (cara is PlanarFace pf)
                                    {
                                        XYZ normal = pf.FaceNormal;
                                        if (normal.Z < 0)
                                        {

                                            Cara_ = cara;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return Cara_;

        }

        

        
    }
}
