using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPI_task42
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand

    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            string pipeInfo = string.Empty;

            var pipes = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_PipeCurves)
                .WhereElementIsNotElementType()
                .Cast<Pipe>()
                .ToList();

            string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "pipes.xlsx");

            using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("Pipes");

                int rowIndex = 0;
                foreach (var pipe in pipes)
                {
                    double externalDiameterValue = 0;
                    double innerDiameterValue = 0;
                    double lengthValue = 0;
                    Parameter externalDiameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                    if (externalDiameter.StorageType == StorageType.Double)
                    {
                        externalDiameterValue = UnitUtils.ConvertFromInternalUnits(externalDiameter.AsDouble(), UnitTypeId.Millimeters);
                    }
                    Parameter innerDiameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM);
                    if (innerDiameter.StorageType == StorageType.Double)
                    {
                        innerDiameterValue = UnitUtils.ConvertFromInternalUnits(innerDiameter.AsDouble(), UnitTypeId.Millimeters);
                    }
                    Parameter length = pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                    if (length.StorageType == StorageType.Double)
                    {
                        lengthValue = UnitUtils.ConvertFromInternalUnits(length.AsDouble(), UnitTypeId.Millimeters);
                    }
                    sheet.SetCellValue(rowIndex, columnIndex: 0, pipe.PipeType.Name);
                    sheet.SetCellValue(rowIndex, columnIndex: 1, externalDiameterValue);
                    sheet.SetCellValue(rowIndex, columnIndex: 2, innerDiameterValue);
                    sheet.SetCellValue(rowIndex, columnIndex: 3, lengthValue);

                    rowIndex++;
                }

                workbook.Write(stream);
                workbook.Close();
            }

            System.Diagnostics.Process.Start(excelPath);

            return Result.Succeeded;
        }
    }
}
