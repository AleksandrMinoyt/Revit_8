using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.DB.Plumbing;

namespace Revit_8
{
    [Transaction(TransactionMode.Manual)]
    public class AddHole : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document arDoc = commandData.Application.ActiveUIDocument.Document;
            Document ovDoc = arDoc.Application.Documents.OfType<Document>().Where(x => x.Title.Contains("ОВ")).FirstOrDefault();

            if(ovDoc==null)
            {
                TaskDialog.Show("Ошибка", "Не найден ОВ файл");
                return Result.Cancelled;
            }


            var holeFamily = new FilteredElementCollector(arDoc)
                   .OfClass(typeof(FamilySymbol))
                   .OfCategory(BuiltInCategory.OST_GenericModel)
                   .OfType<FamilySymbol>()
                   .Where(x => x.FamilyName.Equals("Отверстия"))
                   .FirstOrDefault();

            if (holeFamily == null)
            {
                TaskDialog.Show("Ошибка", "Не найдено семейство отверстий");
                return Result.Cancelled;
            }

            var ducts = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Duct))
                .OfType<Duct>()
                .ToList();


            var pipes = new FilteredElementCollector(ovDoc)
              .OfClass(typeof(Pipe))
              .OfType<Pipe>()
              .ToList();

            //Найдём 3д виды

            var view3d = new FilteredElementCollector(arDoc)
                    .OfClass(typeof(View3D))
                    .Cast<View3D>()
                    .Where(x => x.IsTemplate == false)
                    .FirstOrDefault();


            if (view3d == null)
            {
                TaskDialog.Show("Ошибка", "Нет 3д вида");
                return Result.Cancelled;
            }

            ReferenceIntersector ri = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, view3d);

            Transaction trans = new Transaction(arDoc, "Рассановка отверстий");

            trans.Start();
            if (!holeFamily.IsActive)
                holeFamily.Activate();

            foreach (Duct duct in ducts)
            {
                var curve = (duct.Location as LocationCurve).Curve as Line;
                XYZ point = curve.GetEndPoint(0);
                XYZ direction = curve.Direction;

                List<ReferenceWithContext> intersect = ri.Find(point, direction)
                    .Where(x => x.Proximity <= curve.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer())
                    .ToList();
                foreach (ReferenceWithContext refer in intersect)
                {
                    double proximity = refer.Proximity;
                    Reference reference = refer.GetReference();
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;

                    var level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointHole = point + (direction * proximity);

                    FamilyInstance holeNew = arDoc.Create.NewFamilyInstance(pointHole, holeFamily, wall, level, StructuralType.NonStructural);
                    Parameter width = holeNew.LookupParameter("Ширина");
                    Parameter height = holeNew.LookupParameter("Высота");
                    width.Set(duct.Diameter);
                    height.Set(duct.Diameter);
                }           
              }

            foreach (Pipe pipe in pipes)
            {
                var curve = (pipe.Location as LocationCurve).Curve as Line;
                XYZ point = curve.GetEndPoint(0);
                XYZ direction = curve.Direction;

                List<ReferenceWithContext> intersect = ri.Find(point, direction)
                    .Where(x => x.Proximity <= curve.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer())
                    .ToList();
                foreach (ReferenceWithContext refer in intersect)
                {
                    double proximity = refer.Proximity;
                    Reference reference = refer.GetReference();
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;

                    var level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointHole = point + (direction * proximity);

                    FamilyInstance holeNew = arDoc.Create.NewFamilyInstance(pointHole, holeFamily, wall, level, StructuralType.NonStructural);
                    Parameter width = holeNew.LookupParameter("Ширина");
                    Parameter height = holeNew.LookupParameter("Высота");
                    width.Set(pipe.Diameter);
                    height.Set(pipe.Diameter);
                }
            }


            trans.Commit();
                  

            return Result.Succeeded;
        }
        public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
        {
            public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
            {
                if (ReferenceEquals(x, y)) return true;
                if (ReferenceEquals(null, x)) return false;
                if (ReferenceEquals(null, y)) return false;

                var xReference = x.GetReference();

                var yReference = y.GetReference();

                return xReference.LinkedElementId == yReference.LinkedElementId
                           && xReference.ElementId == yReference.ElementId;
            }

            public int GetHashCode(ReferenceWithContext obj)
            {
                var reference = obj.GetReference();

                unchecked
                {
                    return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
                }
            }
        }

    }
}
