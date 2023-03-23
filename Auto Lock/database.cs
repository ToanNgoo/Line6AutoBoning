using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Collections;

namespace Auto_Lock
{
    public class database_1
    {
        public string constr = @"Provider=Microsoft.Jet.OLEDB.4.0; Data Source = " + Application.StartupPath + @"\Database.mdb";
        public string user;

        public string constr_sever = @"Provider=Microsoft.Jet.OLEDB.4.0; Data Source = " + Application.StartupPath + @"\Database.mdb";
        public OleDbConnection GetConnection()
        {
            OleDbConnection con = new OleDbConnection(constr);
            con.Open();
            return con;
        }

        public OleDbConnection GetConnection_Sever()
        {
            OleDbConnection con = new OleDbConnection(constr_sever);
            con.Open();
            return con;
        }

        //======================== Hàm update database ============================//

        public void update(string sql_update) // hàm thêm data vào sql
        {
            OleDbConnection cnn = new OleDbConnection(constr);
            cnn.Open();
            OleDbCommand cmd = new OleDbCommand(sql_update, cnn);
            cmd.ExecuteNonQuery();
            cnn.Close();
        }

        public void update_sever(string sql)
        {
            OleDbConnection cnn = new OleDbConnection(constr_sever);
            cnn.Open();
            OleDbCommand cmd = new OleDbCommand(sql, cnn);
            cmd.ExecuteNonQuery();
            cnn.Close();
        }
        //======================== Hàm insert database ============================//

        public void insert(string sql_update) // hàm thêm data vào sql
        {
            OleDbConnection cnn = new OleDbConnection(constr);
            cnn.Open();
            OleDbCommand cmd = new OleDbCommand(sql_update, cnn);
            cmd.ExecuteNonQuery();
            cnn.Close();
        }

        //======================== Hàm delete database ============================//

        public void delete(string sql_delete) // hàm xóa data trong sql
        {
            OleDbConnection cnn = new OleDbConnection(constr);
            cnn.Open();
            OleDbCommand cmd = new OleDbCommand(sql_delete, cnn);
            cmd.ExecuteNonQuery();
            cnn.Close();
        }
        public DataTable getData(string str)
        {
            DataTable dt = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(str, constr);
            da.Fill(dt);
            return dt;
        }

        public bool login_part(string user, string pass, string part)
        {
            string str = "select user, password from login where user = '" + user + "' and password = '" + pass + "'and part = '" + part + "'";

            DataTable dt = new DataTable();
            dt = getData(str);

            if (dt.Rows.Count == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool login(string user, string pass)
        {
            string str = "select user, password from login where user = '" + user + "' and password = '" + pass + "'";

            DataTable dt = new DataTable();
            dt = getData(str);

            if (dt.Rows.Count == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool KiemTraKiTu(string s)
        {
            string specialChar = @"\|!#$%&/()=?»«@£§€{}.-;'<>_,";
            for (int i = 0; i < s.Length; i++)
            {
                foreach (var item in specialChar)
                {
                    if (char.IsDigit(s[i]) == false || s.Contains(item) == true)
                    {
                        return true;
                    }
                }
            }

            return false;
        }       
       
        public void insert_infor(DataTable dt)
        {
            OleDbConnection cnn = new OleDbConnection(constr);
            cnn.Open();
            string str = "";
            try
            {
                foreach (DataRow dr in dt.Rows)
                {
                    str = "INSERT INTO Infor VALUES ('" + dr.ItemArray[0] + "', '" + dr.ItemArray[1] + "', '" + dr.ItemArray[2] + "', '" + dr.ItemArray[3] + "')";
                    OleDbCommand cmd = new OleDbCommand(str, cnn);
                    cmd.ExecuteNonQuery();
                }
                cnn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Đã xảy ra lỗi !" + ex.Message, "Chú ý", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public string[] getBarCode(string code)
        {
            int i = 0;
            string st = "Select Code from infor where code = '" + code + "'";
            DataTable dt = new DataTable();
            dt = getData(st);
            i = dt.Rows.Count;
            string[] new_str = new string[i];
            int j = 0;
            foreach (DataRow dr in dt.Rows)
            {
                new_str[j] = dr.ItemArray[0].ToString();
                j++;
            }
            return new_str;
        }
       
    }
}
