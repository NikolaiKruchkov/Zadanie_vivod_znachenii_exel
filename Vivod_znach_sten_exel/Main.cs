using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.Streaming;

namespace Vivod_znach_sten_exel
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            var walls = new FilteredElementCollector(doc)
                .OfClass(typeof(Wall))
                .Cast<Wall>()
                .ToList();
            string wallInfo = string.Empty;
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string exelPath = Path.Combine(desktopPath, "wallInf.xlsx");
            using (FileStream stream = new FileStream(
                exelPath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workBook = new SXSSFWorkbook();
                ISheet sheet = workBook.CreateSheet("Лист1");
                int rowIndex = 0;
                int columnIndex = 0;
                foreach (Wall wall in walls)
                {
                    string wallType = wall.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString();
                    double wallVolume = wall.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED).AsDouble();
                    sheet.SetCellValue(rowIndex, columnIndex, wallType);
                    columnIndex++;
                    sheet.SetCellValue(rowIndex, columnIndex, wallVolume);
                    columnIndex = 0;
                    rowIndex++;                    
                }
                workBook.Write(stream);
                workBook.Close();
            }
            System.Diagnostics.Process.Start(exelPath);
            return Result.Succeeded;
        }
    }
}
