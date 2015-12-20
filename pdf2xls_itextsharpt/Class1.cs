using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System.Reflection;


namespace txt2xls
{
    class Program
    {
        static void xxx(string[] args)
        {
            int row = 1;
            Application app = new Application();
            Workbooks wkbs = app.Workbooks;
            _Workbook wkb = wkbs.Add(@"C:\java_projects\settlement\src\test.xlsx");

            Sheets shs = wkb.Sheets;

            _Worksheet wsh = (_Worksheet)shs.get_Item(1);

            FileStream fs = new FileStream(@"C:\java_projects\settlement\src\RSD20151120-51028B（追保用）.txt", FileMode.Open, FileAccess.Read, FileShare.None);
            StreamReader sr = new StreamReader(fs, Encoding.GetEncoding("GB2312"));
            string str = sr.ReadLine();
            //string str = string.Empty;
            while (str != null)
            {
                if (str == "编号 客户 说明 开始结算日 结束结算日")
                {
                    string[] arr = str.Split(' ');
                    for (int col = 1; col <= arr.Length; col++)
                    {
                        wsh.Cells[row, col] = arr[col - 1];
                    }
                    app.AlertBeforeOverwriting = false;
                    wkb.Save();
                    wkb.SaveAs(@"e:\test.xlsx", Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                    wkb.Close();
                    wkbs.Close();
                    app.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                    app = null;
                    //app.Workbooks.Close();
                    //str = sr.ReadLine();

                    //continue;

                }
                str = sr.ReadLine();
                //switch (str)
                //{
                //    case "":
                //        break;
                //    default:
                //        break;
                //}
            }



            Console.WriteLine(str);
            Console.ReadKey();
        }
    }
}
