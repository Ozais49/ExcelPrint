using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Text;

using NPOI.Util;
using NPOI.SS.Util;
using System.Globalization;

namespace ExcelPrint
{
    internal static class NPOIExtension
    {
        public static string GetFormattedCellValue(this ICell cell, IFormulaEvaluator eval = null)
        {
         

           
            if (cell != null)
            {
                var cellType = cell.CellType;
                switch (cell.CellType)
                {
                    case CellType.String:
                        return cell.StringCellValue;

                    case CellType.Numeric:
                        if (DateUtil.IsCellDateFormatted(cell))
                        {
                            DateTime date = cell.DateCellValue;
                            ICellStyle style = cell.CellStyle;
                            DataFormatter formatter = new DataFormatter();
                            return formatter.FormatCellValue(cell);
                        }
                        else
                        {
                            return cell.NumericCellValue.ToString();
                        }

                    case CellType.Boolean:
                        return cell.BooleanCellValue ? "TRUE" : "FALSE";

                    case CellType.Formula:
                        if (eval != null)
                            return GetFormattedCellValue(eval.EvaluateInCell(cell));
                        else
                            return cell.CellFormula;

                    case CellType.Error:
                        return FormulaError.ForInt(cell.ErrorCellValue).String;
                }
            }
            // null or blank cell, or unknown cell type
            return string.Empty;
        }

        public static string GetTextAlignment(this ICell cell)
        {
            string alignment = string.Empty;

            switch (cell.CellStyle.Alignment)
            {

                case HorizontalAlignment.General:
                    if (cell.CellType == CellType.String)
                    {
                        alignment = "Left";
                    }
                    else if (cell.CellType == CellType.Boolean)
                    {
                        alignment = "Center";
                    }
                    else
                    {
                        alignment = "Right";
                    }
                    break;
                //case HorizontalAlignment.Left:
                //case HorizontalAlignment.Right:
                //case HorizontalAlignment.Center:
                //    return cell.CellStyle.Alignment.ToString();
                default:
                    alignment = cell.CellStyle.Alignment.ToString();
                    break;
            }
            return $"text-align:{alignment};";

        }

        public static string GetFontStyle(this ICell cell,IWorkbook workbook)
        {
            StringBuilder fontStyleBuilder = new StringBuilder();
            var font = cell.CellStyle.GetFont(workbook);
            var fontSize = font.FontHeightInPoints;
            var fontName = font.FontName;
            var fontcolor = ((XSSFFont)font).GetXSSFColor().RGB;
            //font-family:{fontName};
            //font-size:{fontSize}px;
            fontStyleBuilder.Append($"font-size:{fontSize}pt;font-family:{fontName};color:rgb({fontcolor[0]},{fontcolor[1]},{fontcolor[2]});");
            if (font.IsBold)
            {
                fontStyleBuilder.Append($"font-weight:bold;");
            }
            if (font.IsItalic)
            {
                fontStyleBuilder.Append("font-style:italic;");
            }
            if (font.IsStrikeout)
            {
                fontStyleBuilder.Append("text-decoration: line-through;");
            }

            return fontStyleBuilder.ToString();
        }

        public static string GetBackgroundColor(this ICell cell)
        {
            string background = "background-color:rgb(255,255,255);";
            if (cell.CellStyle.FillPattern == FillPattern.SolidForeground)
            {
                byte[] cellBackground = ((XSSFColor)cell.CellStyle.FillForegroundColorColor).RGB;
                background = $"background-color:rgb({cellBackground[0]},{cellBackground[1]},{cellBackground[2]});";
            }
            return background;
        }
        public static string RemoveEndingSemiColon(this string stringValue)
        {
            return stringValue.TrimEnd(';');
        }
    }
}
