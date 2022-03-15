using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB.Structure;

namespace RevitAPIUI10CreateModel
{
    [Transaction(TransactionMode.Manual)]
    public class CreateModel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var doc = commandData.Application.ActiveUIDocument.Document;

            GetWalls(commandData);

            List<Level> levels = GetLevels(commandData);
            List<Wall> walls = GetWallTypes(commandData);

            Transaction ts1 = new Transaction(doc, "Построение дверей");
            {
                ts1.Start();
                AddDoor(commandData, levels.Where(x => x.Name.Equals("Уровень 1")).FirstOrDefault(), walls[0]);
                ts1.Commit();
            }

            Transaction ts2 = new Transaction(doc, "Построение окон");
            {
                ts2.Start();
                foreach (Wall wall in walls)
                {
                    if (!wall.Equals(walls[0]))
                    AddWindow(commandData, levels.Where(x => x.Name.Equals("Уровень 1")).FirstOrDefault(), wall);
                    
                }
                ts2.Commit();
            }
            
            return Result.Succeeded;
        }

        public static List<Level> GetLevels(ExternalCommandData commandData)
        {
            var doc = commandData.Application.ActiveUIDocument.Document;

            List<Level> levels = new FilteredElementCollector(doc)
                                                       .OfClass(typeof(Level))
                                                       .Cast<Level>()
                                                       .ToList();
           

            return levels;
        }

        public static List<XYZ> GetPoints()
        {
            double width = UnitUtils.ConvertToInternalUnits(10000, UnitTypeId.Millimeters);
            double depth = UnitUtils.ConvertToInternalUnits(5000, UnitTypeId.Millimeters);

            double dx = width / 2;
            double dy = depth / 2;

            List<XYZ> points = new List<XYZ>();
            points.Add(new XYZ(-dx, -dy, 0));
            points.Add(new XYZ(dx, -dy, 0));
            points.Add(new XYZ(dx, dy, 0));
            points.Add(new XYZ(-dx, dy, 0));
            points.Add(new XYZ(-dx, -dy, 0));

            return points;
        }

        public static List<Wall> GetWalls(ExternalCommandData commandData)
        {
            var doc = commandData.Application.ActiveUIDocument.Document;
            List<Level> levels = GetLevels(commandData);
            List<XYZ> points = GetPoints();
            List<Wall> walls = new List<Wall>();

            Level level1 = levels
              .Where(x => x.Name.Equals("Уровень 1"))
              .FirstOrDefault();
            Level level2 = levels
                .Where(x => x.Name.Equals("Уровень 2"))
                .FirstOrDefault();

            Transaction ts = new Transaction(doc, "Построение стен");
            {
                ts.Start();
                for (int i = 0; i < 4; i++)
                {
                    Line line = Line.CreateBound(points[i], points[i + 1]);
                    Wall wall = Wall.Create(doc, line, level1.Id, false);
                    walls.Add(wall);
                    wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(level2.Id);
                }
                ts.Commit();
            }

            return walls;
        }

        private static void AddDoor(ExternalCommandData commandData, Level level, Wall wall)
        {
            var doc = commandData.Application.ActiveUIDocument.Document;
            List<Level> levels = GetLevels(commandData);

            FamilySymbol doorType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0915 x 2134 мм"))
                .Where(x => x.FamilyName.Equals("Одиночные-Щитовые"))
                .FirstOrDefault();

            LocationCurve hostCurve = wall.Location as LocationCurve;
            XYZ point1 = hostCurve.Curve.GetEndPoint(0);
            XYZ point2 = hostCurve.Curve.GetEndPoint(1);
            XYZ midPoint = (point1 + point2) / 2;
            
            if (!doorType.IsActive)
                doorType.Activate();

            doc.Create.NewFamilyInstance(midPoint, doorType, wall, levels.Where(x => x.Name.Equals("Уровень 1")).FirstOrDefault(), StructuralType.NonStructural);
        }

        private static void AddWindow(ExternalCommandData commandData, Level level, Wall wall)
        {
            var doc = commandData.Application.ActiveUIDocument.Document;

            List<Level> levels = GetLevels(commandData);
            List<Wall> walls = GetWallTypes(commandData);

            FamilySymbol windowType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Windows)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0610 x 1220 мм"))
                .Where(x => x.FamilyName.Equals("Фиксированные"))
                .FirstOrDefault();

            LocationCurve hostCurve = wall.Location as LocationCurve;
            XYZ point1 = hostCurve.Curve.GetEndPoint(0);
            XYZ point2 = hostCurve.Curve.GetEndPoint(1);
            XYZ midPoint1 = (point1 + point2) / 2;
            XYZ midPoint2 = GetElementCenter(wall);
            XYZ offset = (midPoint1 + midPoint2) / 2;

            if (!windowType.IsActive)
                windowType.Activate();

            doc.Create.NewFamilyInstance(offset, windowType, wall, levels.Where(x => x.Name.Equals("Уровень 1")).FirstOrDefault(), StructuralType.NonStructural);
        }

        public static XYZ GetElementCenter(Element element)
        {
            BoundingBoxXYZ bounding = element.get_BoundingBox(null);
            return (bounding.Min + bounding.Max) / 2;
        }

        public static List<Wall> GetWallTypes(ExternalCommandData commandData)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var walllist = new FilteredElementCollector(doc)
                                                       .OfClass(typeof(Wall))
                                                       .Cast<Wall>()
                                                       .ToList();

            return walllist;
        }
    }
}
