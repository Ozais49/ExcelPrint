using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPrint
{
    public class PDF
    {
        public void CreatePdf()
        {
            PdfDocument pdf = new PdfDocument();
            pdf.Info.Title = "Created With PDFSharp";
            PdfPage pdfpage = pdf.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(pdfpage);
            gfx.DrawString("Hello, World!", font, XBrushes.Black,
                      new XRect(0, 0, pdfpage.Width, pdfpage.Height),
                      XStringFormat.Center);




        }
    }
}
