using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;

using log4net;
using iTextSharp.text.pdf;
using iTextSharp.text;
using iTextSharp.text.pdf.codec;
using iTextSharp.text.pdf.parser;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

[assembly: log4net.Config.XmlConfigurator(Watch = true)]
namespace pdf2xls_itextsharpt
{
    class Program
    {

        static void Main(string[] args)
        {
            log4net.ILog log = log4net.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
            log.Info("hello error");
            //Document doc = new Document();
            //PdfWriter.GetInstance(doc, new System.IO.FileStream(@"d:\\settlement\\bills\\RSD20151120-51028B（追保用）.pdf", FileMode.Open));

            //PdfDocument pdfdoc = new PdfDocument("settlefile");
            string FileName = @"RSD20151120-51028B（追保用）";
            PdfReader reader = new PdfReader(FileName + ".pdf");

            string text = string.Empty;
            int n = reader.NumberOfPages;
            //byte[] b = new byte[0];
            for(int i = 1; i <= n;i++)
            {
                text += PdfTextExtractor.GetTextFromPage(reader, i);
               text += "\r\n";
                //reader.GetPageContent(i);
                //StringBuilder sb = new StringBuilder();

            }
            FileStream fs = new FileStream(FileName + ".txt", FileMode.Create);
            StreamWriter sw = new StreamWriter(fs, Encoding.GetEncoding("GB2312"));
            sw.Write(text);
            sw.Close();
            fs.Close();

            FileStream rfs = new FileStream(FileName + ".txt", FileMode.Open);
            StreamReader sr = new StreamReader(rfs, Encoding.GetEncoding("GB2312"));
            Application app = new Application();
            Workbooks wkbs = app.Workbooks;
           // _Workbook wkb = wkbs.Add(FileName + ".xlsx");
            _Workbook wkb = wkbs.Add("test.xlsx");

            Sheets shs = wkb.Sheets;
            

            _Worksheet wsh = (_Worksheet)shs.get_Item(1);

            String line = sr.ReadLine();
            
            while(line != null)
            {
                wsh.Cells[1, 1] = line;
                break;
            }
            wkb.SaveAs(@".\test.xlsx", Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            wkb.Close();
            wkbs.Close();
            app.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            app = null;

            reader.Close();

            Console.ReadKey();
        }
    }
}
