using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;

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
            //string FileName = @"RSD20151120-51028B（追保用）";
            string FileName = @"RSD20151120-51028D（追保用）";
            
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
            wsh.Name = FileName;
            String line = sr.ReadLine();
            int rowcounter = 1;
            bool  dropline = false;
            int six_check = 0;
            int serven_check = 0;
            int eight_check = 0;
            int nine_check = 0;
            int ten_check = 0;
            int eleven_check = 0;
            int twelve_check = 0;
            while(line != null)
            {
                if (line == "3. 资产状况变化图")
                {
                    dropline = true;
                }
                if (line == "6. 今日出入金-明细")
                {
                    dropline = false;
                }

                if (!dropline)
                {
                    if (line.StartsWith("Page") || IsDate(line) )//|| six_check || serven_check || eight_check || nine_check || ten_check || eleven_check || twelve_check) 
                    {
                        line = sr.ReadLine();
                        continue;
                    }
                    else
                    {
                        if (line == "编号 日期 币种 入金 出金 备注")
                        {
                            six_check++;
                            if(six_check > 1)
                            {
                                line = sr.ReadLine();
                                continue;
                            }
                        }
                        if (line == "编号 客户号 合约 币种 方向 价格 手数 发生时间 手续费")
                        {
                            serven_check++;
                            if (serven_check > 1)
                            {
                                line = sr.ReadLine();
                                continue;
                            }
                        }
                        if (line == "编号 客户号 结算日 合约 币种 买入价格 买入单号 卖出价格 卖出单号 手数 类型 逐日盯市盈亏 逐笔对冲盈亏")
                        {
                            eight_check++;
                            if (eight_check > 1)
                            {
                                line = sr.ReadLine();
                                continue;
                            }
                        }
                        if (line == "编号 客户号 交易日 合约 币种 单号 方向 成交价格 上次结算价格 本次结算价格 手数 逐日盯市盈亏 逐笔对冲盈亏")
                        {
                            nine_check++;
                            if (nine_check > 1)
                            {
                                line = sr.ReadLine();
                                continue;
                            }
                        }
                        if (line == "编号 币种 合约 合约到期日 买持仓 买均价 卖持仓 卖均价 昨日结算价 今日结算价 逐日盯市盈亏 逐笔对冲盈亏")
                        {
                            ten_check++;
                            if (ten_check > 1)
                            {
                                line = sr.ReadLine();
                                continue;
                            }
                        }
                        if (line == "编号 合约 到期日 币种 买入价格 买入单号 卖出价格 卖出单号 手数 逐笔对冲盈亏")
                        {
                            eleven_check ++;
                            if (eleven_check > 1)
                            {
                                line = sr.ReadLine();
                                continue;
                            }
                        }
                        if (line == "编号 合约 到期日 币种 买入价格 买入单号 卖出价格 卖出单号 手数 逐笔对冲盈亏")
                        {
                            twelve_check ++;
                            if (twelve_check > 1)
                            {
                                line = sr.ReadLine();
                                continue;
                            }
                        }
                        string[] columns = line.Split(' ');
                        for (int columncounter = 1; columncounter <= columns.Length; columncounter++)
                        {
                            wsh.Cells[rowcounter, columncounter] = columns[columncounter - 1];
                        }
                        rowcounter++;
                        line = sr.ReadLine();
                    }
                }
                else
                {
                    line = sr.ReadLine();
                }
            }
            wkb.SaveAs(@"c:\test.xlsx", Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            wkb.Close();
            wkbs.Close();
            app.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            app = null;

            reader.Close();
            log.Debug("the end!");
            Console.ReadKey();
        }

        public static bool IsDate(string StrSource)
        {
            return Regex.IsMatch(StrSource, @"^((((1[6-9]|[2-9]\d)\d{2})-(0?[13578]|1[02])-(0?[1-9]|[12]\d|3[01]))|(((1[6-9]|[2-9]\d)\d{2})-(0?[13456789]|1[012])-(0?[1-9]|[12]\d|30))|(((1[6-9]|[2-9]\d)\d{2})-0?2-(0?[1-9]|1\d|2[0-9]))|(((1[6-9]|[2-9]\d)(0[48]|[2468][048]|[13579][26])|((16|[2468][048]|[3579][26])00))-0?2-29-))$");
        }

    }
}
