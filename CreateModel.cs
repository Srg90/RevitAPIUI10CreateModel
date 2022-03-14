using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPIUI10CreateModel
{
    [Transaction(TransactionMode.Manual)]
    public class CreateModel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //UIApplication uiapp = commandData.Application;
            //UIDocument uidoc = uiapp.ActiveUIDocument;
            //Document doc = uidoc.Document;

            GetWalls(commandData);           

            return Result.Succeeded;
        }

        public static List<Level> GetLevels(ExternalCommandData commandData)
        {
            var doc = commandData.Application.ActiveUIDocument.Document;

            List<Level> levels = new FilteredElementCollector(doc)
                                                       .OfClass(typeof(Level))
                                                       .Cast<Level>()
                                                       .ToList();
            Level level1 = levels
               .Where(x => x.Name.Equals("Уровень 1"))
               .FirstOrDefault();
            Level level2 = levels
                .Where(x => x.Name.Equals("Уровень 2"))
                .FirstOrDefault();

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

            Transaction ts = new Transaction(doc, "Построение стен");
            {
                ts.Start();
                for (int i = 0; i < 4; i++)
                {
                    Line line = Line.CreateBound(points[i], points[i + 1]);
                    Wall wall = Wall.Create(doc, line, levels.Where(x=>x.Name.Equals("Уровень 1")).FirstOrDefault().Id, false);
                    walls.Add(wall);
                    wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(levels.Where(x => x.Name.Equals("Уровень 2")).FirstOrDefault().Id);
                }
                ts.Commit();
            }

            return walls;
        }
    }
}
