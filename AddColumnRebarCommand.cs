using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace add
{
    [Transaction(TransactionMode.Manual)]
    public class AddColumnRebarCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                Reference r = uidoc.Selection.PickObject(ObjectType.Element, new ColumnSelectionFilter(), "Selecciona una columna estructural");
                var col = doc.GetElement(r) as FamilyInstance;

                if (col == null)
                {
                    TaskDialog.Show("Error", "El elemento seleccionado no es una columna.");
                    return Result.Failed;
                }
                // The user can adjust these parameters as needed, or you can implement a UI to input them
                //==========================================}
                double spacingEndZones = MnToFt(100); // 100 mm
                double maxSpacingCenter = MnToFt(150); // 150 mm
                
                int nBotton = 8; // Number of stirrups in the bottom confining zone
                int nTop = 8; // Number of stirrups in the top confining zone
                //=========================================
                double minSpaceEnds = MnToFt(50);
                
                // Ther first stirups of face floor or beam
                double minConf = MnToFt(500); // 500 mm

                RebarBarType tieBarType = new FilteredElementCollector(doc)
                    .OfClass(typeof(RebarBarType))
                    .Cast<RebarBarType>()
                    .FirstOrDefault();


                using (Transaction tran = new Transaction(doc, "Crear estribos en columna"))
                {
                    tran.Start();
                    ColumnRebarService.CreateRebarZoned(
                        doc, 
                        col,
                        tieBarType,
                        maxSpacingCenter: maxSpacingCenter,
                        spacingEndZones: spacingEndZones,
                        nBottom: nBotton,
                        nTop: nTop,
                        minConfLength: minConf
                        );
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

    public class ColumnSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem.Category != null && elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralColumns;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
    public static class ColumnRebarService
    {
        public static void CreateRebarZoned(
            Document doc,
            FamilyInstance column,
            RebarBarType tieBarType,
            double maxSpacingCenter,
            double spacingEndZones,
            int nBottom,
            int nTop,
            double minConfLength
        )
        {
            Transform t = column.GetTransform();
            Transform inv = t.Inverse;

            (XYZ minL, XYZ maxL) = GetlocalExtents(column, inv);

            double columnHeight = maxL.Z - minL.Z;
            if (columnHeight <= 0.0) throw new InvalidOperationException("La altura de la columna debe ser mayor que cero");

            double conf = Math.Max(columnHeight / 6.0, minConfLength);

            double midLen = columnHeight - 2 * conf;
            if (midLen < 0.0) midLen = 0;

            

            double cover = UnitUtils.ConvertToInternalUnits(40, UnitTypeId.Millimeters);

            double tieDia = tieBarType.get_Parameter(BuiltInParameter.REBAR_BAR_DIAMETER).AsDouble();
            double off = cover + 0.5 * tieDia;

            double x1 = minL.X + off;
            double x2 = maxL.X - off;
            double y1 = minL.Y + off;
            double y2 = maxL.Y - off;

            if (x2 <= x1 || y2 <= y1)
                throw new InvalidOperationException("Sección demasiado pequeña para el recubrimiento/diámetro seleccionado");

            double firstTie = UnitUtils.ConvertToInternalUnits(50, UnitTypeId.Millimeters);

            double z0 = minL.Z /*+ firstTie - 0.5 * tieDia*/;

            List <Curve> profileLocal = CreateRectProfileLocal(x1, x2, y1, y2, z0);
            List<Curve> profileWorld = profileLocal.Select(c => c.CreateTransformed(t)).ToList();

            XYZ normal = t.BasisZ.Normalize();

            //double spacingEndZones = AddColumnRebarCommand.MnToFt(100); // 0.10 m
            //int nBottomCalc = CalcBarsBySpacing(conf - firstTie, spacingEndZones);
            //int nTopCalc = CalcBarsBySpacing(conf - firstTie, spacingEndZones);

            CreateTieSet(doc, column, tieBarType, profileWorld, normal,
            startOffsetFromBase: firstTie,
            runLength: conf - firstTie,
            layout: LayoutKind.FixedNumber,
            spacing: spacingEndZones,
            number: nBottom);

            double bottomEnd = firstTie + (nBottom - 1) * spacingEndZones;
            double topStart = columnHeight - firstTie - (nTop - 1) * spacingEndZones;

            double midStart = bottomEnd /*+ maxSpacingCenter*/;
            double midAvailable = topStart - midStart;

            if (midAvailable > 1e-6)
            {
                CreateTieSet(doc, column, tieBarType, profileWorld, normal,
                    startOffsetFromBase: midStart,
                    runLength: midAvailable,
                    layout: LayoutKind.FixedSpacing,
                    spacing: maxSpacingCenter,
                    number: 0);
            }

            CreateTieSet(doc, column, tieBarType, profileWorld, normal,
                startOffsetFromBase: columnHeight - firstTie - (nTop - 1) * spacingEndZones,
                runLength: conf - firstTie,
                layout: LayoutKind.FixedNumber,
                spacing: spacingEndZones,
                number: nTop);
        }

        private enum LayoutKind
        {
            FixedNumber,
            FixedSpacing
        }
        private static int CalcBarsBySpacing(double runLength, double spacing)
        {
            if (runLength <= 0) return 2;

            int n = (int)Math.Floor(runLength / spacing) + 1;

            if (n < 2) n = 2;

            return n;
        }

        private static void CreateTieSet(
            Document doc,
            FamilyInstance host,
            RebarBarType barType,
            List<Curve> profileWorld,
            XYZ normal,
            double startOffsetFromBase,
            double runLength,
            LayoutKind layout,
            double spacing,
            int number
            )
        {
            Transform tr = Transform.CreateTranslation(normal * startOffsetFromBase);
            IList<Curve> movedProfile = profileWorld.Select(c => c.CreateTransformed(tr)).ToList();

            Rebar rebar = Rebar.CreateFromCurves(
                doc, 
                RebarStyle.StirrupTie,
                barType,
                null,
                null,
                host,
                normal,
                movedProfile,
                RebarHookOrientation.Left,
                RebarHookOrientation.Right,
                true,
                true
                );

            RebarShapeDrivenAccessor acc = rebar.GetShapeDrivenAccessor();

            if (layout == LayoutKind.FixedNumber)
            {
                if (number < 2) number = 2;

                acc.SetLayoutAsNumberWithSpacing(
                    number,
                    spacing,
                    true,
                    true,
                    true
                );
            }
            else
            {
                acc.SetLayoutAsMaximumSpacing(
                    spacing,
                    runLength,
                    true,
                    false,
                    false

                );
            }
        }


        private static List<Curve> CreateRectProfileLocal(double x1, double x2, double y1, double y2, double z)
        {
            XYZ p1 = new XYZ(x1, y1, z);
            XYZ p2 = new XYZ(x2, y1, z);
            XYZ p3 = new XYZ(x2, y2, z);
            XYZ p4 = new XYZ(x1, y2, z);

            return new List<Curve>()
            {
                Line.CreateBound(p1, p2),
                Line.CreateBound(p2, p3),
                Line.CreateBound(p3, p4),
                Line.CreateBound(p4, p1)
            };
        }

        private static (XYZ min, XYZ max) GetlocalExtents(FamilyInstance column, Transform inv)
        {
            BoundingBoxXYZ bb = column.get_BoundingBox(null);
            if (bb == null)
                throw new InvalidOperationException("No se pudo obtener el bounding box de la columna");

            var vertices = new List<XYZ>()
            {
                new XYZ(bb.Min.X, bb.Min.Y, bb.Min.Z),
                new XYZ(bb.Min.X, bb.Min.Y, bb.Max.Z),
                new XYZ(bb.Min.X, bb.Max.Y, bb.Min.Z),
                new XYZ(bb.Min.X, bb.Max.Y, bb.Max.Z),
                new XYZ(bb.Max.X, bb.Min.Y, bb.Min.Z),
                new XYZ(bb.Max.X, bb.Min.Y, bb.Max.Z),
                new XYZ(bb.Max.X, bb.Max.Y, bb.Min.Z),
                new XYZ(bb.Max.X, bb.Max.Y, bb.Max.Z),

            };

            var local = vertices.Select(v => inv.OfPoint(v)).ToList();
            double minX = local.Min(p => p.X);
            double minY = local.Min(p => p.Y);
            double minZ = local.Min(p => p.Z);

            double maxX = local.Max(p => p.X);
            double maxY = local.Max(p => p.Y);
            double maxZ = local.Max(p => p.Z);

            return (new XYZ(minX, minY, minZ), new XYZ(maxX, maxY, maxZ));

        }
    }
}
