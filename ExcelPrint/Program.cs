using GemBox.Spreadsheet;
using NPOI.HSSF.Record;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using PdfSharp;
using PdfSharp.Pdf;
using Spire.Xls.Core.Interfaces;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using TheArtOfDev.HtmlRenderer.PdfSharp;
//using TheArtOfDev.HtmlRenderer.PdfSharp;

namespace ExcelPrint
{
    class Program
    {
    
        static void Main(string[] args)
        {


            // var excelfilePath = "D:\\ExcelReportDocuementFolder.xlsx";
            var excelfilePath = @"D:\GTS Documents\Doorstroomlijst_2020-05-05T13-36-08.942+05-45.xlsx";//Doorstroomlijst_2020-05-05T14-20-10.218+05-45.xlsx";
            var printer = "Microsoft Print to PDF";
            //ExcelToHtml excelToHtml = new ExcelToHtml();
            //var htmlTable = excelToHtml.GetHtMLNPOI(excelfilePath);
            //File.WriteAllText(@"D:\test1.html", htmlTable);
            //PdfDocument pdf = PdfGenerator.GeneratePdf(htmlTable, PageSize.A4,margin:0);
            //pdf.Save(@"D:\\document.pdf");
            //Console.WriteLine(htmlTable);



            //XSSFWorkbook workbook;
            //using (FileStream file = new FileStream(excelfilePath, FileMode.Open, FileAccess.Read))
            //{
            //    workbook = new XSSFWorkbook(file);
            //}
            //var test = workbook.NumberOfSheets;
            //var worksheet = workbook.GetSheetAt(0);
            //MemoryStream ms = new MemoryStream();

            ExcelPrintInterop op = new ExcelPrintInterop();
            op.PrintExcel(excelfilePath,printer);

            //Spireprin print = new Spireprin();
            //print.PrintExcel( excelfilePath,"Microsoft Print to PDF");

            //ExcelPrintGemBox excelprint = new ExcelPrintGemBox();
            //excelprint.Print(excelfilePath, "Microsoft Print to PDF");



            //XlsxLoadOptions loadOptions = new XlsxLoadOptions();
            //ExcelFile excel = ExcelFile.Load("D:\\ExcelReportDocuementFolder.xlsx", loadOptions);

            //PdfSaveOptions saveOptions = new PdfSaveOptions();
            //excel.Save("D:\\ExcelReportDocuementFolder.pdf", saveOptions);
        }

        //public static string GetHtMLClosedXML(string excelfilePath)
        //{
        //    var workbook = new XLWorkbook(excelfilePath);
        //    var numberofSheets = workbook.Worksheets.Count;
        //    var worksheet = workbook.Worksheet(0);
        //    var firstUsedRows = "";
        //    StringBuilder builder = new StringBuilder();
        //    for (int i = 0; i <= worksheet.LastRowUsed().Value; i++)
        //    {

        //        var row = worksheet.GetRow(i);
        //        if (row.Hidden.HasValue)
        //        {
        //            if (!row.Hidden.Value)
        //            {
        //                builder.AppendLine("<tr>");
        //                foreach (var cell in row)
        //                {
        //                    builder.AppendLine($"<td>{cell}</td>");
        //                }
        //            }
        //        }
        //    }

        //    return "";

        //}


    

    }
}
