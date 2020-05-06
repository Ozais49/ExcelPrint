using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPrint
{
    public class ExcelToHtml
    {
        private bool IsXLSX { get; set; }
        private IWorkbook Workbook { get; set; }
        public string tableStyle = "border-collapse: collapse;";//font-family: helvetica, arial, sans-serif;
        public string GetHtMLNPOI(string excelfilePath)
        {
            Workbook = GetWorkbook(excelfilePath);
            var numberofSheets = Workbook.NumberOfSheets;
            var worksheet = Workbook.GetSheetAt(0);
            IFormulaEvaluator evaluator = Workbook.GetCreationHelper().CreateFormulaEvaluator();
            var rowbreaks=worksheet.RowBreaks;
            
            StringBuilder builder = new StringBuilder();
           
            for (int i = 0; i <= worksheet.LastRowNum; i++)
            {

                var row = worksheet.GetRow(i);
                if (row == null)
                {
                    builder.AppendLine("<tr></tr>");
                    continue;
                }
                if (row.Hidden.HasValue)
                {
                    if (!row.Hidden.Value)
                    {
                        builder.AppendLine("<tr>");
                        
                        foreach (var cell in row)
                        {
                            var columnWidth = worksheet.GetColumnWidthInPixels(cell.ColumnIndex);
                            var style = GetCellStyle(cell,columnWidth);
                            builder.AppendLine($"<td style=\"{style}\">{cell.GetFormattedCellValue(evaluator)}</td>");
                        }
                        builder.AppendLine("</tr>");
                    }
                }
            }

            return $"<table style=\"{tableStyle}\">{builder}</table>";

        }

        private IWorkbook GetWorkbook(string excelfilePath)
        {
           // FileInfo fileInfo = new FileInfo(excelfilePath);

            using (FileStream file = new FileStream(excelfilePath, FileMode.Open, FileAccess.Read))
            {
                //if (string.Equals(fileInfo.Extension, ".XLSX", StringComparison.OrdinalIgnoreCase))
                //{
                    IsXLSX = true;
                    return new XSSFWorkbook(file);
                //}
                //else
                //{
                //    IsXLSX = false;
                //    return new HSSFWorkbook(file);
                //}

            }
        }
        private string GetCellStyle(ICell cell,float cellwidth)
        {
        
            var cellStyle = cell.CellStyle;
            var alignment = cell.GetTextAlignment().RemoveEndingSemiColon();
            var font = cell.GetFontStyle(Workbook).RemoveEndingSemiColon();
            var backgroundColor = cell.GetBackgroundColor().RemoveEndingSemiColon();
            return $"width:{cellwidth}px;{alignment};{font};{backgroundColor};";
            
        }
        //public string GetColorStyle(ICell cell)
        //{
        //    if(cell.CellStyle.color)
        //    return "";

        //}
       

    }
}
