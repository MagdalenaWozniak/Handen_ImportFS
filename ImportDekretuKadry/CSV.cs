using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportFS
{
    public class CSV
    {
      
            public  string KontoWn { get; set; }
            public string KwotaWn { get; set; }
            public string KontoMa { get; set; }
            public string KwotaMa { get; set; }
            public  string Opis { get; set; }
            
           
            public static CSV FromCsv(string csvLine)
            {
            string[] values = csvLine.Split(';');
                CSV csvlist = new CSV();
                csvlist.KontoWn = values[0];
                csvlist.KwotaWn = values[1];
                csvlist.KontoMa = values[2];
                csvlist.KwotaMa = values[3];

            csvlist.Opis = RemoveSpecialCharacters(values[4]);// values[4];//Replace("'", "l").Replace("˘", "o").Replace("ž", "s").Replace("Ť", "c");
            return csvlist;
            }

        private static string RemoveSpecialCharacters(string str)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char c in str)
            {
                if ((c >= '0' && c <= '9') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '.' || c == '_')
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }

    }
}
