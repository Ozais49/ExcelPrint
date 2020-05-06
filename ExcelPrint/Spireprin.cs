using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Xls;

namespace ExcelPrint
{
   public class Spireprin
    {  

        public void PrintExcel(string excelFilePath,string printerName)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(excelFilePath);
            //Get the first sheet and save to csv format file  

           
            PrintDocument print = workbook.PrintDocument;
            print.PrintController = new StandardPrintController();
            print.PrinterSettings.PrinterName = printerName;
            print.Print();
        }
      
    }
}
