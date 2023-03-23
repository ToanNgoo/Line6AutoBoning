using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;

using System.IO;


using Excel = Microsoft.Office.Interop.Excel;
namespace Auto_Lock
{
    class ClsExcel
    {
        public void Export_CSV(DataTable dt, string path, bool check, string ColumnName)
        {
            StringBuilder sb = new StringBuilder();

            if (check == false)
            {
                sb.Append(ColumnName);
            }
            foreach (DataRow dr in dt.Rows)
            {
                foreach (DataColumn dc in dt.Columns)
                    sb.Append(FormatCSV(dr[dc.ColumnName].ToString()) + ",");
                sb.Remove(sb.Length - 1, 1);
                sb.AppendLine();
            }
            if (check == false)
            {               
                File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
            }
            else
                File.AppendAllText(path, sb.ToString(), Encoding.UTF8);
        }

        public void Export_CS(DataTable dt, string path, string ColumnName)
        {
            StringBuilder sb = new StringBuilder();           
            sb.Append(ColumnName);
            
            foreach (DataRow dr in dt.Rows)
            {
                foreach (DataColumn dc in dt.Columns)
                    sb.Append(FormatCSV(dr[dc.ColumnName].ToString()) + ",");
                sb.Remove(sb.Length - 1, 1);
                sb.AppendLine();
            }            
                File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
            
        }
        public static string FormatCSV(string input)
        {
            try
            {
                if (input == null)
                    return string.Empty;

                bool containsQuote = false;
                bool containsComma = false;
                int len = input.Length;
                for (int i = 0; i < len && (containsComma == false || containsQuote == false); i++)
                {
                    char ch = input[i];
                    if (ch == '"')
                        containsQuote = true;
                    else if (ch == ',')
                        containsComma = true;
                }

                if (containsQuote && containsComma)
                    input = input.Replace("\"", "\"\"");

                if (containsComma)
                    return "\"" + input + "\"";
                else
                    return input;
            }
            catch
            {
                throw;
            }
        }

    }

}
