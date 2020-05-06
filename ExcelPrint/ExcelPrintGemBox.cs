using GemBox.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPrint
{
   public class ExcelPrintGemBox
    {
        public void Print(string excelFile,string printerName)
        {
            //used library is paid one and need to set  key before using
            //consists of limitations
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            XlsxLoadOptions loadOptions = new XlsxLoadOptions();
            ExcelFile excel = ExcelFile.Load(excelFile, loadOptions);
            excel.Print(printerName);
            //save as pdf if print not work
            //PdfSaveOptions saveOptions = new PdfSaveOptions();
            //excel.Save(excelFile, saveOptions);
        }
      
    }
}
