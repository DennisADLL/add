using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;

namespace add
{
    [Transaction(TransactionMode.Manual)]
    public class AddBeamRebarCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                Reference r = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    new BeamSelectionFilter(),
                    "Selecciona una viga estructural");

                var beam = doc.GetElement(r) as FamilyInstance;

                if (beam == null)
                {
                    TaskDialog.Show("Error", "El elemento seleccionado no es una viga.");
                    return Result.Failed;
                }

                double spacingEndZones = MmToFt(100);   // 100 mm
                double maxSpacingCenter = MmToFt(150); // 150 mm

                int nLeft = 5;   // zona confinada inicial
                int nRight = 5;  // zona confinada final

                double minConf = MmToFt(500); // longitud mínima de confinamiento
                double cover = MmToFt(40);    // recubrimiento

                RebarBarType tieBarType = new FilteredElementCollector(doc)
                    .OfClass(typeof(RebarBarType))
                    .Cast<RebarBarType>()
                    .FirstOrDefault();

                if (tieBarType == null)
                    throw new InvalidOperationException("No se encontró ningún RebarBarType.");

                RebarHookType hookType = new FilteredElementCollector(doc)
                    .OfClass(typeof(RebarHookType))
                    .Cast<RebarHookType>()
                    .FirstOrDefault(h => h.Name == "Sísmico de estribo/tirante - 135°.");

                using (Transaction tran = new Transaction(doc, "Crear estribos en viga"))
                {
                    tran.Start();

                    BeamRebarService.CreateRebarZoned(
                        doc,
                        beam,
                        tieBarType,
                        hookType,
                        cover,
                        maxSpacingCenter,
                        spacingEndZones,
                        nLeft,
                        nRight,
                        minConf
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

        public static double MmToFt(double mm)
        {
            return UnitUtils.ConvertToInternalUnits(mm, UnitTypeId.Millimeters);
        }
    }

    public class BeamSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem.Category != null &&
                   elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }

    public static class BeamRebarService
    {
        public static void CreateRebarZoned(
            Document doc,
            FamilyInstance beam,
            RebarBarType tieBarType,
            RebarHookType hookType,
            double cover,
            double maxSpacingCenter,
            double spacingEndZones,
            int nLeft,
            int nRight,
            double minConfLength
        )
        {
            Transform t = beam.GetTransform();
            Transform inv = t.Inverse;

            (XYZ minL, XYZ maxL) = GetLocalExtents(beam, inv);

            double beamLength = maxL.X - minL.X;
            if (beamLength <= 0.0)
                throw new InvalidOperationException("La longitud de la viga debe ser mayor que cero.");

            double conf = Math.Max(beamLength / 6.0, minConfLength);

            double midLen = beamLength - 2 * conf;
            if (midLen < 0.0) midLen = 0.0;

            double tieDia = tieBarType.get_Parameter(BuiltInParameter.REBAR_BAR_DIAMETER).AsDouble();

            double off = cover + 0.5 * tieDia;

            double y1 = minL.Y + off;
            double y2 = maxL.Y - off;
            double z1 = minL.Z + off;
            double z2 = maxL.Z - off;

            if (y2 <= y1 || z2 <= z1)
                throw new InvalidOperationException("La sección de la viga es demasiado pequeña para el recubrimiento y diámetro seleccionados.");

            double firstTie = AddBeamRebarCommand.MmToFt(50);

            double x0 = minL.X;

            List<Curve> profileLocal = CreateRectProfileLocalYZ(x0, y1, y2, z1, z2, tieDia);
            List<Curve> profileWorld = profileLocal.Select(c => c.CreateTransformed(t)).ToList();

            XYZ normal = t.BasisX.Normalize();

            // Zona izquierda
            CreateTieSet(
                doc,
                beam,
                tieBarType,
                hookType,
                profileWorld,
                normal,
                startOffsetFromStart: firstTie,
                runLength: conf - firstTie,
                layout: LayoutKind.FixedNumber,
                spacing: spacingEndZones,
                number: nLeft
            );

            double leftEnd = firstTie + (nLeft - 1) * spacingEndZones;
            double rightStart = beamLength - firstTie - (nRight - 1) * spacingEndZones;

            double midStart = leftEnd;
            double midAvailable = rightStart - midStart;

            if (midAvailable > 1e-6)
            {
                CreateTieSet(
                    doc,
                    beam,
                    tieBarType,
                    hookType,
                    profileWorld,
                    normal,
                    startOffsetFromStart: midStart,
                    runLength: midAvailable,
                    layout: LayoutKind.FixedSpacing,
                    spacing: maxSpacingCenter,
                    number: 0
                );
            }

            // Zona derecha
            CreateTieSet(
                doc,
                beam,
                tieBarType,
                hookType,
                profileWorld,
                normal,
                startOffsetFromStart: beamLength - firstTie - (nRight - 1) * spacingEndZones,
                runLength: conf - firstTie,
                layout: LayoutKind.FixedNumber,
                spacing: spacingEndZones,
                number: nRight
            );
        }

        private enum LayoutKind
        {
            FixedNumber,
            FixedSpacing
        }

        private static void CreateTieSet(
            Document doc,
            FamilyInstance host,
            RebarBarType barType,
            RebarHookType hookType,
            List<Curve> profileWorld,
            XYZ normal,
            double startOffsetFromStart,
            double runLength,
            LayoutKind layout,
            double spacing,
            int number
        )
        {
            Transform tr = Transform.CreateTranslation(normal * startOffsetFromStart);
            IList<Curve> movedProfile = profileWorld.Select(c => c.CreateTransformed(tr)).ToList();

            Rebar rebar = Rebar.CreateFromCurves(
                doc,
                RebarStyle.StirrupTie,
                barType,
                hookType,
                hookType,
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
                    true,
                    true
                );
                acc.SetLayoutAsMaximumSpacing(
                    spacing,
                    runLength,
                    true,
                    false,
                    false
                );


            }
        }

        private static List<Curve> CreateRectProfileLocalYZ(
            double x,
            double y1,
            double y2,
            double z1,
            double z2,
            double tieDia
        )
        {
            // Perfil abierto tipo estribo en plano YZ
            XYZ p1 = new XYZ(x, y1 - 0.5 * tieDia, z1);
            XYZ p2 = new XYZ(x, y2, z1);
            XYZ p3 = new XYZ(x, y2, z2);
            XYZ p4 = new XYZ(x, y1, z2);
            XYZ p5 = new XYZ(x, y1, z1 - 0.5 * tieDia);

            return new List<Curve>()
            {
                Line.CreateBound(p1, p2),
                Line.CreateBound(p2, p3),
                Line.CreateBound(p3, p4),
                Line.CreateBound(p4, p5)
            };
        }

        private static (XYZ min, XYZ max) GetLocalExtents(FamilyInstance beam, Transform inv)
        {
            BoundingBoxXYZ bb = beam.get_BoundingBox(null);
            if (bb == null)
                throw new InvalidOperationException("No se pudo obtener el bounding box de la viga.");

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