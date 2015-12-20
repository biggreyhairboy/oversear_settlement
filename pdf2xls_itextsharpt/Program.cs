using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

using iTextSharp.text.pdf;
using iTextSharp.text;
using iTextSharp.text.pdf.codec;
using iTextSharp.text.pdf.parser;


namespace pdf2xls_itextsharpt
{
    class Program
    {
        static void Main(string[] args)
        {
            //Document doc = new Document();
            //PdfWriter.GetInstance(doc, new System.IO.FileStream(@"d:\\settlement\\bills\\RSD20151120-51028B（追保用）.pdf", FileMode.Open));

            //PdfDocument pdfdoc = new PdfDocument("settlefile");
            string FileName = @"c:\settlement\bills\RSD20151120-51028B（追保用）";
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
            reader.Close();

            Console.ReadKey();

            

        }
    }
}
