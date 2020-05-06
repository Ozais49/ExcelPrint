using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using RawPrint;

namespace ExcelPrint
{
    public class ExcelPrintInterop
    {
        public void PrintExcel(string excelFile, string printerName)
        {
            Application application = new Application();
            Workbook workbook = application.Workbooks.Open(excelFile,
                                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Worksheet ws = (Worksheet)workbook.Worksheets[1];

            try
            {

                var range=ws.UsedRange;
                
                SetPrinterName(printerName, application);
                ws.PageSetup.PaperSize = XlPaperSize.xlPaperA4;
                ws.PageSetup.Orientation = XlPageOrientation.xlLandscape;
                ws.PageSetup.FitToPagesWide = 1;
                ws.PageSetup.FitToPagesTall = false;
                ws.PageSetup.Zoom =false;

                //void PrintOut(object From, object To, object Copies, object Preview, 
                //object ActivePrinter, object PrintToFile, object Collate, object PrToFileName);

                // Print out 1 copy to the default printer:
                ws.PrintOut(
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                // Cleanup:

              
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Marshal.FinalReleaseComObject(ws);

                workbook.Close(false, Type.Missing, Type.Missing);
                Marshal.FinalReleaseComObject(workbook);

                application.Quit();
                Marshal.FinalReleaseComObject(application);
                application = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }


        }

        private void SetPrinterName(string printerName,Application application)
        {
           string port=string.Empty;
            try
            {
                
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows NT\CurrentVersion\Devices"))
                {
                    if (key != null)
                    {
                        object value = key.GetValue(printerName);
                        if (value != null)
                        {
                            string[] values = value.ToString().Split(',');
                            if (values.Length >= 2) port = values[1];
                        }
                    }
                }

                if (!application.ActivePrinter.StartsWith(printerName))
                {
                    // Get current concatenation string ('on' in enlgish, 'op' in dutch, etc..)
                    var split = application.ActivePrinter.Split(' ');
                    if (split.Length >= 3)
                    {
                        application.ActivePrinter = String.Format("{0} {1} {2}",
                            printerName,
                            split[split.Length - 2],
                            port);
                    }
                   
                }

            }
            catch (Exception e)
            {
               //return e.Message;
            }
          //  return port;

        }

    }
}
