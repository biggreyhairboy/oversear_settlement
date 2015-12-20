using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace pdf2xls_itextsharpt
{
    class Utility
    {
        public static string [] SplitRows(string str, char splitter = ' ')
        {
            return str.Split(splitter);
        }
    }
}
