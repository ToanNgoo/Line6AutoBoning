using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.IO.Ports;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;

namespace Auto_Lock
{
    public partial class Form1 : Form
    {
        OptionDefine.clsCheckTrungInformation checkcod = new OptionDefine.clsCheckTrungInformation();
        List<string> listInfor = new List<string>();
        clsPLC PLC = new clsPLC();
        //clsPLC PLC_bond = new clsPLC();
        clsPLC_1 PLC_bond = new clsPLC_1();
        ClsExcel Excel = new ClsExcel();
        clsScanner Sr1000;
        database_1 ClsDtb = new database_1();
        int PCM_No1 = 0;
        int PCM_No2 = 0;
        int PCM_No3 = 0;
        int status_1 = 0;
        int status_2 = 0;
        int status_3 = 0;
        string start = string.Empty;
        string linkFile = string.Empty;
        string model_test = string.Empty;
        string test = string.Empty;
        bool memory = true;
        string error = string.Empty;
        string Jig_End = string.Empty;
        string _memory = string.Empty;
        string CH_Enable = string.Empty;
        string CH_1 = string.Empty;
        string CH_2 = string.Empty;
        string CH_3 = string.Empty;
        DataTable dt = new DataTable();
        string[] _model_name = new string[3];
        bool kq = false;
        string _File = String.Empty;
        string result = String.Empty;
        string[] OK_NG = new string[1];
        string[] barcode = new string[1];
        string[] CH = new string[1];
        string[] time_Test = new string[1];
        DataTable TableName = new DataTable();
        string model_Run = string.Empty;
        //----------------------------------//
        string Ready_bond = string.Empty;
        int End_check = 0;
        string End_bond = string.Empty;
        string vTri_bond_1 = string.Empty;
        string vTri_bond_2 = string.Empty;
        string vTri_bond_3 = string.Empty;
        string codeSR = string.Empty;

        bool check_Ready = true;
        bool check_end = true;
        bool vTri1 = true;
        bool vTri2 = true;
        bool vTri3 = true;

        string read_Code1 = string.Empty;
        string read_Code2 = string.Empty;
        string read_Code3 = string.Empty;
        int read_position = 0;
        int Error = 0;
        int vtriStatus, vtriCode, vtriTime, vtriCH;
        string server_PBA = string.Empty;

        int PCM_Date = 0;

        string link_Server = string.Empty;
        string link_PC_Boxing = string.Empty;
        public Form1()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
            openFileDialog1.Filter = "CSV files (*.csv)|*.csv|Excel |*.xlsx|Excel 2003 | *.xls";
            openFileDialog2.Filter = "CSV files (*.csv)|*.csv|Excel |*.xlsx|Excel 2003 | *.xls";

            btn_browseMaster.Enabled = false;
            btn_browseActual.Enabled = false;
            btn_saveLinkMaster.Enabled = false;
            btn_excel.Enabled = false;
            txt_fileName.ReadOnly = true;
            txt_fileNameActual.ReadOnly = true;
            chb_changeMaster.Enabled = false;

        }

        public void Sr1000_Datareceived(object sender, SerialDataReceivedEventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                lb_check.Text = "";
                model_test = cb_Model.Text;
                
                if (cb_Model.Text.ToUpper().Contains("MAIN"))
                {                   
                    model_Run = "MAIN";
                }

                if (cb_Model.Text.ToUpper().Contains("CELL"))
                {                   
                    model_Run = "CELL";
                }

                if (cb_Model.Text.ToUpper().Contains("SUB"))
                {                    
                    model_Run = "SUB";
                }

                string month = DateTime.Now.Month.ToString("00");
                linkFile = File.ReadAllText(@Application.StartupPath + "\\LinkFile.txt");

                if (chb_changeMaster.Checked == true)
                {
                    string a = "";
                }
                else
                {
                    string link = string.Empty;
                    link = System.IO.File.ReadAllText("Link.txt"); // link log File
                    _File = getFile(cb_Model.Text, link);
                    lb_checkDate.Text = "";
                    lb_checkDate.BackColor = Color.Transparent;

                    lb_pass.BackColor = Color.Transparent;
                    lb_pass.Text = "";

                    txt_fileName.Text = @Application.StartupPath + "\\File Master\\" + model_Run + ".csv";
                    txt_fileNameActual.Text = _File.ToString();

                    //DateTime dateTime = DateTime.Now;
                    //DateTime startDateTime = new DateTime(dateTime.Year, dateTime.Month, dateTime.Day, 7, 45, 0);
                    //DateTime endDateTime = new DateTime(dateTime.Year, dateTime.Month, dateTime.Day, 19, 30, 0);

                    //DateTime dateTimeNight = DateTime.Now.AddDays(1);
                    //DateTime startDateTimeNight = new DateTime(dateTime.Year, dateTime.Month, dateTime.Day, 19, 45, 0);
                    //DateTime endDateTimeNight = new DateTime(dateTimeNight.Year, dateTimeNight.Month, dateTimeNight.Day, 7, 30, 0);

                    ////DateTime lastWriteTime = Convert.ToDateTime(File.GetLastWriteTime(_File).ToString("00/00/0000 HH:00:ss", CultureInfo.InvariantCulture));
                    //DateTime lastWriteTime = DateTime.Parse(File.GetLastWriteTime(_File).ToString());
                    //lb_checkDate.Text = "OK";
                    //lb_checkDate.BackColor = Color.Green;

                    //if ((lastWriteTime > startDateTimeNight && lastWriteTime < endDateTimeNight)
                    //    || (lastWriteTime > startDateTime && lastWriteTime < endDateTime))
                    //{
                    //    lb_checkDate.Text = "OK";
                    //    lb_checkDate.BackColor = Color.Green;                      
                    //}
                    //else
                    //{
                    //    lb_checkDate.Text = "NG";
                    //    lb_checkDate.BackColor = Color.Red;
                    //    MessageBox.Show("Hãy test bản đầu tiên đầu mỗi Shift trước khi kiểm tra thông số function! \nFile test mới nhất lúc : " + File.GetLastWriteTime(_File), "Chú ý", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    //    File.AppendAllText(@Application.StartupPath + "\\Result\\log.txt", "\r\n"
                    //            + DateTime.Now.ToString() + "\r\nName : " + tb_User.Text + "\r Model "
                    //            + cb_Model.Text + " Check link so sanh NG \r\nFile test mới nhất lúc : " + File.GetLastWriteTime(_File) + "");
                    //}
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //SR1000.Open();
            lb_check.Text = "";
            if (!Directory.Exists(@Application.StartupPath + "\\File Master"))
            {
                Directory.CreateDirectory(@Application.StartupPath + "\\File Master");
            }

            if (!File.Exists(@Application.StartupPath + "\\Config.txt"))
            {
                File.WriteAllText(@Application.StartupPath + "\\Config.txt", "");
            }

            string[] data = File.ReadAllLines(@Application.StartupPath + "\\Config.txt");
            for (int i = 0; i < data.Length; i++)
            {
                if (data[i].Contains(lb_StartRead.Text))
                {
                    string[] pos = data[i].Split('|');
                    tb_StartRead.Text = pos[1];
                }

                if (data[i].Contains(lb_modelRun.Text))
                {
                    string[] pos = data[i].Split('|');
                    tb_modelRun.Text = pos[1];
                }

                if (data[i].Contains(lb_StartBond.Text))
                {
                    string[] pos = data[i].Split('|');
                    tb_Startbond.Text = pos[1];
                }

                if (data[i].Contains(lb_EndBond.Text))
                {
                    string[] pos = data[i].Split('|');
                    tb_EndBond.Text = pos[1];
                }

                if (data[i].Contains(lb_linkdata.Text))
                {
                    string[] pos = data[i].Split('|');
                    tb_Enable.Text = pos[1];
                }

                if (data[i].Contains(lb_Keo1.Text))
                {
                    string[] pos = data[i].Split('|');
                    tb_Keo1.Text = pos[1];
                }

                if (data[i].Contains(lb_Keo2.Text))
                {
                    string[] pos = data[i].Split('|');
                    tb_Keo2.Text = pos[1];
                }

                if (data[i].Contains(lb_Keo3.Text))
                {
                    string[] pos = data[i].Split('|');
                    tb_Keo3.Text = pos[1];
                }

                if (data[i].Contains(lb_EndRead.Text))
                {
                    string[] pos = data[i].Split('|');
                    tb_EndRead.Text = pos[1];
                }

                if (data[i].Contains(lb_PLCFuntion.Text))
                {
                    string[] pos = data[i].Split('|');
                    tb_IP1.Text = pos[1];
                }

                if (data[i].Contains(lb_PLCBonding.Text))
                {
                    string[] pos = data[i].Split('|');
                    tb_IP2.Text = pos[1];
                }

                if (data[i].Contains(lb_stopbond.Text))
                {
                    string[] pos = data[i].Split('|');
                    tb_stopbond.Text = pos[1];
                }
                if (data[i].Contains(lb_Main.Text))
                {
                    string[] pos = data[i].Split('|');
                    tb_Main.Text = pos[1];
                    //tb_model1.Text = pos[1];
                }
                if (data[i].Contains(lb_Cell.Text))
                {
                    string[] pos = data[i].Split('|');
                    tb_Cell.Text = pos[1];
                    //tb_model2.Text = pos[1];
                }


                if (data[i].Contains(lb_DPos.Text))
                {
                    string[] pos = data[i].Split('|');
                    tb_DPos.Text = pos[1];
                }

                if (data[i].Contains(lb_pos1.Text))
                {
                    string[] pos = data[i].Split('|');
                    tb_Pos1.Text = pos[1];
                }

                if (data[i].Contains(lb_pos2.Text))
                {
                    string[] pos = data[i].Split('|');
                    tb_Pos2.Text = pos[1];
                }

                if (data[i].Contains(lb_pos3.Text))
                {
                    string[] pos = data[i].Split('|');
                    tb_Pos3.Text = pos[1];
                }

                if (data[i].Contains(lb_ComScanner.Text))
                {
                    string[] pos = data[i].Split('|');
                    tb_Scanner.Text = pos[1];
                }
            }
            tb_Enable.Enabled = false;
            btn_save2.Enabled = false;
            //tb_Reset.Enabled = false;
            tb_IP1.Enabled = false;
            tb_IP2.Enabled = false;
            tb_Scanner.Enabled = false;
            tb_StartRead.Enabled = false;
            tb_EndRead.Enabled = false;
            tb_DPos.Enabled = false;
            tb_Pos1.Enabled = false;
            tb_Pos2.Enabled = false;
            tb_Pos3.Enabled = false;
            tb_modelRun.Enabled = false;
            tb_StartRead.Enabled = false;
            tb_EndBond.Enabled = false;
            tb_stopbond.Enabled = false;
            tb_CurPos.Enabled = false;
            tb_Main.Enabled = false;
            tb_Cell.Enabled = false;
            tb_Keo1.Enabled = false;
            tb_Keo2.Enabled = false;
            tb_Keo3.Enabled = false;
            button6.Enabled = false;
            button4.Enabled = false;
            btn_Save.Enabled = false;
            btn_change.Enabled = false;
            tb_Startbond.Enabled = false;
            Sr1000 = new clsScanner(this);
            Sr1000.COMnum = tb_Scanner.Text;
            Sr1000.ketnoi(bt_CntScan);
            PLC.thietlap(tb_IP1);
            PLC.ketnoi(lb_PLC, bt_CntPlcFct);
            PLC_bond.thietlap(tb_IP2);
            PLC_bond.ketnoi(lb_PLCbond, bt_CntPlcBon);
            string[] array = File.ReadAllLines(@Application.StartupPath + "\\Model.txt");
            cb_Model.Items.AddRange(array);
            _model_name = File.ReadAllLines(@Application.StartupPath + "\\Model.txt");

            if (!File.Exists(@Application.StartupPath + "\\Server.log"))
            {
                File.Create(@Application.StartupPath + "\\Server.log");
            }
            try
            {
                //string[] data = null;
                FileStream FS = new FileStream(@Application.StartupPath + "\\Server.log", FileMode.Open);
                StreamReader SR = new StreamReader(FS);
                while (SR.EndOfStream == false)
                {
                    link_Server = SR.ReadLine();
                }
                SR.Close();
                FS.Close();
            }
            catch (Exception)
            {
                MessageBox.Show("Không kết nối được đến server, liên hệ PE giải quyết");
            }

            try
            {
                //string[] data = null;
                FileStream FS = new FileStream(@Application.StartupPath + "\\Boxing.log", FileMode.Open);
                StreamReader SR = new StreamReader(FS);
                while (SR.EndOfStream == false)
                {
                    link_PC_Boxing = SR.ReadLine();
                }
                SR.Close();
                FS.Close();
            }
            catch (Exception)
            {
                MessageBox.Show("Không kết nối được đến máy log, liên hệ PE giải quyết");
            }
        }



        private void button2_Click(object sender, EventArgs e)
        {
            if (cb_Model.Text == "")
            {
                MessageBox.Show("Bạn chưa chọn model, hãy chọn model đang chạy");
            }
            else
            {
                timer1.Enabled = true;
                try
                {
                    dgv_actual.Columns.Clear();
                    dgv_master.Columns.Clear();
                    kq = kiemTra(@Application.StartupPath + "\\File Master\\" + model_Run + ".csv", _File);
                    if (kq == true)
                    {
                        lb_check.BackColor = Color.Green;
                        lb_check.Text = "OK";
                        MessageBox.Show("Master thông số OK. File test mới nhất lúc : " + File.GetLastWriteTime(_File));
                        File.AppendAllText(@Application.StartupPath + "\\Result\\log.txt", "\r\n"
                        + DateTime.Now.ToString() + "\r\nName : " + tb_User.Text + "\r Check Master "
                        + cb_Model.Text + "  OK\r\nFile test moi nhat luc : " + File.GetLastWriteTime(_File));
                    }
                    else
                    {
                        lb_pass.BackColor = Color.Red;
                        lb_pass.Text = "NG";
                        result = "Result check : NG";
                        MessageBox.Show("Master thông số NG. Hãy kiểm tra lại! \nFile test mới nhất lúc : " + File.GetLastWriteTime(_File), "Chú ý", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        File.AppendAllText(@Application.StartupPath + "\\Result\\log.txt", "\r\n"
                            + DateTime.Now.ToString() + "\r\nName : " + tb_User.Text + "\r Check Master "
                            + cb_Model.Text + "  NG\r\nFile test moi nhat luc : " + File.GetLastWriteTime(_File));
                    }

                }
                catch (Exception)
                {
                    MessageBox.Show("Không tìm thấy file check sheet thông số. Hãy kiểm tra lại");
                    lb_check.BackColor = Color.Red;
                    lb_check.Text = "NG";
                }

                //try
                //{
                //    Create_File();
                //}
                //catch (Exception ee)
                //{
                //    MessageBox.Show(ee.ToString());
                //}
            }
            //ClsDtb.delete("Delete *from infor");
        }

        // Xử lí log File
        public void modelTest(int vTri_barcode, int vTri_Ok_NG, int vTri_Time, int vTri_CH)
        {
            PCM_No1 = 0;
            PCM_No2 = 0;
            PCM_No3 = 0;
            string _File = getFile(model_test, linkFile);
            dt = ReadCsvFile(_File);

            string[] mang1 = new string[dt.Rows.Count];
            string[] mang1_1 = new string[dt.Rows.Count];
            string[] mang1_2 = new string[dt.Rows.Count];
            string[] testTime = new string[dt.Rows.Count];
            string[] barcodeArray = new string[dt.Rows.Count];
            int j = 0;

            foreach (DataRow dr in dt.Rows)
            {
                if (j < dt.Rows.Count - 1)
                {
                    if (j < 10)
                    {
                        mang1[j] = dr.ItemArray[0].ToString().Substring(1, 1);
                    }
                    else
                    {
                        mang1[j] = dr.ItemArray[0].ToString().Substring(1, 2);
                    }
                    mang1_1[j] = dr.ItemArray[vTri_Ok_NG].ToString();
                    mang1_2[j] = dr.ItemArray[vTri_CH].ToString();
                    barcodeArray[j] = dr.ItemArray[vTri_barcode].ToString();
                    testTime[j] = dr.ItemArray[vTri_Time].ToString();
                    j++;
                }
            }
            if (mang1_2[dt.Rows.Count - 2] == "3")
            {
                PCM_No1 = dt.Rows.Count - 3;
                PCM_No2 = dt.Rows.Count - 2;
                PCM_No3 = dt.Rows.Count - 1;
                string[] _testTime = new string[3];
                string[] barcode = new string[3];
                string[] OK_NG = new string[3];
                barcode[0] = barcodeArray[PCM_No1 - 1];
                barcode[1] = barcodeArray[PCM_No2 - 1];
                barcode[2] = barcodeArray[PCM_No3 - 1];
                string[] chanel = { "1", "2", "3" };
                if (mang1_1[PCM_No1 - 1] == "OK")
                {
                    status_1 = 1;
                    OK_NG[0] = "OK";
                    _testTime[0] = testTime[PCM_No1 - 1];
                }
                else
                {
                    status_1 = 0;
                    OK_NG[0] = "NG";
                    _testTime[0] = testTime[PCM_No1 - 1];
                }

                if (mang1_1[PCM_No2 - 1] == "OK")
                {
                    status_2 = 1;
                    OK_NG[1] = "OK";
                    _testTime[1] = testTime[PCM_No2 - 1];
                }
                else
                {
                    status_2 = 0;
                    OK_NG[1] = "NG";
                    _testTime[1] = testTime[PCM_No2 - 1];
                }

                if (mang1_1[PCM_No3 - 1] == "OK")
                {
                    status_3 = 1;
                    OK_NG[2] = "OK";
                    _testTime[2] = testTime[PCM_No3 - 1];
                }
                else
                {
                    status_3 = 0;
                    OK_NG[2] = "NG";
                    _testTime[2] = testTime[PCM_No3 - 1];
                }

                TableName = Table_new(OK_NG, barcode, _testTime, chanel);

            }
            else if (mang1_2[dt.Rows.Count - 2] == "2")
            {
                PCM_No1 = dt.Rows.Count - 2;
                PCM_No2 = dt.Rows.Count - 1;
                PCM_No3 = dt.Rows.Count;
                status_3 = 0;
                string[] barcode = new string[2];
                string[] OK_NG = new string[2];
                string[] _testTime = new string[2];
                barcode[0] = barcodeArray[PCM_No1 - 1];
                barcode[1] = barcodeArray[PCM_No2 - 1];
                string[] chanel = { "1", "2" };
                if (mang1_1[PCM_No1 - 1] == "OK")
                {
                    status_1 = 1;
                    OK_NG[0] = "OK";
                    _testTime[0] = testTime[PCM_No1 - 1];
                }
                else
                {
                    status_1 = 0;
                    OK_NG[0] = "NG";
                    _testTime[0] = testTime[PCM_No1 - 1];
                }

                if (mang1_1[PCM_No2 - 1] == "OK")
                {
                    status_2 = 1;
                    OK_NG[1] = "OK";
                    _testTime[1] = testTime[PCM_No2 - 1];
                }
                else
                {
                    status_2 = 0;
                    OK_NG[1] = "NG";
                    _testTime[1] = testTime[PCM_No2 - 1];
                }

                TableName = Table_new(OK_NG, barcode, _testTime, chanel);
            }
            else if (mang1_2[dt.Rows.Count - 2] == "1")
            {
                PCM_No1 = dt.Rows.Count - 1;
                PCM_No2 = dt.Rows.Count;
                PCM_No3 = dt.Rows.Count + 1;
                status_3 = 0;
                status_2 = 0;
                string[] _testTime = new string[1];
                string[] barcode = new string[1];
                string[] OK_NG = new string[1];
                barcode[0] = barcodeArray[PCM_No1 - 1];
                string[] chanel = { "1" };
                if (mang1_1[PCM_No1 - 1] == "OK")
                {
                    status_1 = 1;
                    OK_NG[0] = "OK";
                    _testTime[0] = testTime[PCM_No1 - 1];
                }
                else
                {
                    status_1 = 0;
                    OK_NG[0] = "NG";
                    _testTime[0] = testTime[PCM_No1 - 1];
                }
                TableName = Table_new(OK_NG, barcode, _testTime, chanel);

            }

            if (model_Run == "CELL" || model_Run == "MAIN")
            {
                ClsDtb.insert_infor(TableName);
            }
        }
        public string getFile(string model, string link)
        {
            try
            {
                int month = int.Parse(DateTime.Now.Month.ToString());
                string linkFile = string.Empty;
                linkFile = link + "\\" + model + "\\" + DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString("00");
                string[] file = Directory.GetFiles(linkFile); // link đến vị trí fileLog

                DateTime[] dt = new DateTime[file.Length];
                for (int i = 0; i < file.Length; i++)
                {
                    dt[i] = File.GetLastWriteTime(file[i]);
                }
                DateTime maxTime = dt[0];
                int vTri = 0;
                for (int i = 1; i < dt.Length; i++)
                {
                    int soSanh = DateTime.Compare(dt[i], maxTime);
                    if (soSanh > 0)
                    {
                        maxTime = dt[i];
                        vTri = i;
                    }
                }
                return file[vTri];
            }
            catch (Exception)
            {
                return "";
            }

        }
        public void output()
        {
            if (CH_Enable == "3")
            {
                if (status_1 == 1)
                {
                    PLC.Writeplc("Y48", 0);
                    PLC.Writeplc("Y4F", 1);
                }
                else
                {
                    PLC.Writeplc("Y48", 1);
                    PLC.Writeplc("Y4F", 0);

                }
                if (status_2 == 1)
                {
                    PLC.Writeplc("Y4A", 0);
                    PLC.Writeplc("Y49", 1);

                }
                else
                {
                    PLC.Writeplc("Y4A", 1);
                    PLC.Writeplc("Y49", 0);

                }

                if (status_3 == 1)
                {
                    PLC.Writeplc("Y4E", 0);
                    PLC.Writeplc("Y4B", 1);
                }
                else
                {
                    PLC.Writeplc("Y4E", 1);
                    PLC.Writeplc("Y4B", 0);
                }
            }
            else if (CH_Enable == "2")
            {
                PLC.Writeplc("Y4E", 0);
                PLC.Writeplc("Y4B", 0);
                if (status_1 == 1)
                {
                    PLC.Writeplc("Y48", 0);
                    PLC.Writeplc("Y4F", 1);

                }
                else
                {
                    PLC.Writeplc("Y48", 1);
                    PLC.Writeplc("Y4F", 0);

                }
                if (status_2 == 1)
                {
                    PLC.Writeplc("Y4A", 0);
                    PLC.Writeplc("Y49", 1);

                }
                else
                {
                    PLC.Writeplc("Y4A", 1);
                    PLC.Writeplc("Y49", 0);

                }
            }
            else if (CH_Enable == "1")
            {
                PLC.Writeplc("Y4E", 0);
                PLC.Writeplc("Y4B", 0);
                PLC.Writeplc("Y4A", 0);
                PLC.Writeplc("Y49", 0);
                if (status_1 == 1)
                {
                    PLC.Writeplc("Y48", 0);
                    PLC.Writeplc("Y4F", 1);

                }
                else
                {
                    PLC.Writeplc("Y48", 1);
                    PLC.Writeplc("Y4F", 0);

                }

            }
        }

        //Xử lí File CSV
        public DataTable ReadCsvFile(string path)
        {
            DataTable dtb = new DataTable();
            string Fulltext;
            using (StreamReader sr = new StreamReader(path))
            {
                while (!sr.EndOfStream)
                {
                    Fulltext = sr.ReadToEnd().ToString();
                    string[] rows = Fulltext.Split('\r');

                    for (int i = 5; i < rows.Length; i++)
                    {
                        string[] rowValues = rows[i].Split(',');

                        if (i == 5)
                        {
                            for (int j = 0; j < rowValues.Count(); j++)
                            {
                                dtb.Columns.Add(rowValues[j]);                                
                            }
                        }
                        else if (i == 6 || i == 7 || i == 8 || i == 9)
                        {
                            ;
                        }
                        else
                        {
                            DataRow dr = dtb.NewRow();
                            for (int k = 0; k < rowValues.Count(); k++)
                            {                                
                                dr[k] = rowValues[k].ToString();
                               
                            }
                            dtb.Rows.Add(dr);
                        }

                    }
                }
            }
            return dtb;
        }
        public DataTable ReadCsvFile(string path, int lineStart)
        {
            DataTable dtb = new DataTable();
            string Fulltext;
            using (StreamReader sr = new StreamReader(path))
            {
                while (!sr.EndOfStream)
                {
                    Fulltext = sr.ReadToEnd().ToString();
                    string[] rows = Fulltext.Split('\r');

                    for (int i = lineStart; i < rows.Length; i++)
                    {
                        string[] rowValues = rows[i].Split(',');

                        if (i == lineStart)
                        {
                            for (int j = 0; j < rowValues.Count(); j++)
                            {
                                dtb.Columns.Add(rowValues[j]);
                            }
                        }
                        else
                        {
                            DataRow dr = dtb.NewRow();
                            for (int k = 0; k < rowValues.Count(); k++)
                            {
                                dr[k] = rowValues[k];

                                if (dr[k].ToString().Contains(@""""))
                                {
                                    dr[k] = "";
                                }
                            }
                            dtb.Rows.Add(dr);
                        }
                    }
                }
            }
            return dtb;
        }
        public DataTable ReadCsvFile(string path, int lineStart, int lineEnd)
        {
            DataTable dtb = new DataTable();
            string Fulltext;
            using (StreamReader sr = new StreamReader(path))
            {
                while (!sr.EndOfStream)
                {
                    Fulltext = sr.ReadToEnd().ToString();
                    string[] rows = Fulltext.Split('\r');

                    for (int i = lineStart; i < lineEnd; i++)
                    {
                        string[] rowValues = rows[i].Split(',');

                        if (i == 5)
                        {
                            for (int j = 0; j < rowValues.Count(); j++)
                            {
                                dtb.Columns.Add(rowValues[j]);

                                if (rowValues[j].ToString() == "Total Result")
                                {
                                    vtriStatus = j;
                                }

                                if (rowValues[j].ToString() == "Barcode(SN)")
                                {
                                    vtriCode = j;
                                }

                                if (rowValues[j].ToString() == "Test Time")
                                {
                                    vtriTime = j;
                                }

                                if (rowValues[j].ToString() == "Test Channel")
                                {
                                    vtriCH = j;
                                }                               

                                if (rowValues[j].ToString() == "Read Master PCM Date" || rowValues[j].ToString() == "Manufacture Date")
                                {
                                    PCM_Date = j;
                                }

                            }
                        }
                        else
                        {
                            DataRow dr = dtb.NewRow();
                            for (int k = 0; k < rowValues.Count(); k++)
                            {
                                dr[k] = rowValues[k];

                                if (dr[k].ToString().Contains(@""""))
                                {
                                    dr[k] = "";
                                }
                            }
                            dtb.Rows.Add(dr);
                        }
                    }
                }
            }
            return dtb;
        }

        //Xử lí barcode
        public void XuLyCode()
        {
            try
            {
                if (PLC_bond.readplc(tb_Keo1.Text) == "0" && PLC_bond.readplc(tb_Keo2.Text) == "0" && PLC_bond.readplc(tb_Keo3.Text) == "0")
                {
                    PLC_bond.Writeplc(tb_stopbond.Text, 1);
                    Thread.Sleep(200);
                    PLC_bond.Writeplc(tb_stopbond.Text, 0);
                    step = 1;
                }
                else
                {
                    PLC_bond.Writeplc(tb_Startbond.Text, 1);
                    Thread.Sleep(200);
                    PLC_bond.Writeplc(tb_Startbond.Text, 0);
                    step = 5;
                }
            }
            catch (Exception)
            {
                ;
            }
        }

        #region So sánh barcode
        public void CompareCode(TextBox tb)
        {
            string[] array = new string[4];
            DataTable dt;
            try
            {
                dt = ClsDtb.getData("Select *from Infor where barcode = '" + tb.Text.Substring(0, 14) + "'");
                if (dt.Rows.Count == 1)
                {
                    if (dt.Rows[0].ItemArray[0].ToString() == "OK")
                    {
                        if (PLC_bond.readplc(tb_modelRun.Text) == tb_Cell.Text) // model Cell
                        {
                            if (dt.Rows[0].ItemArray[3].ToString() == "1")
                            {
                                PLC_bond.Writeplc(tb_Keo1.Text, 1);
                                Thread.Sleep(200);
                            }
                            if (dt.Rows[0].ItemArray[3].ToString() == "2")
                            {
                                PLC_bond.Writeplc(tb_Keo2.Text, 1);
                                Thread.Sleep(200);
                            }
                            if (dt.Rows[0].ItemArray[3].ToString() == "3")
                            {
                                PLC_bond.Writeplc(tb_Keo3.Text, 1);
                                Thread.Sleep(200);
                            }
                        }
                        else if (PLC_bond.readplc(tb_modelRun.Text) == tb_Main.Text) // model Main
                        {
                            PLC_bond.Writeplc(tb_Keo1.Text, 0);
                            Thread.Sleep(100);

                            if (dt.Rows[0].ItemArray[3].ToString() == "1")
                            {
                                PLC_bond.Writeplc(tb_Keo2.Text, 1);
                                Thread.Sleep(200);
                            }
                            if (dt.Rows[0].ItemArray[3].ToString() == "2")
                            {
                                PLC_bond.Writeplc(tb_Keo3.Text, 1);
                                Thread.Sleep(200);
                            }
                        }
                    }
                    else
                    {
                        if (PLC_bond.readplc(tb_modelRun.Text) == tb_Cell.Text)
                        {
                            if (dt.Rows[0].ItemArray[3].ToString() == "1")
                            {
                                PLC_bond.Writeplc(tb_Keo1.Text, 0);
                                Thread.Sleep(200);
                            }
                            if (dt.Rows[0].ItemArray[3].ToString() == "2")
                            {
                                PLC_bond.Writeplc(tb_Keo2.Text, 0);
                                Thread.Sleep(200);
                            }
                            if (dt.Rows[0].ItemArray[3].ToString() == "3")
                            {
                                PLC_bond.Writeplc(tb_Keo3.Text, 0);
                                Thread.Sleep(200);
                            }
                        }
                        else if (PLC_bond.readplc(tb_modelRun.Text) == tb_Main.Text)
                        {
                            if (dt.Rows[0].ItemArray[3].ToString() == "1")
                            {
                                PLC_bond.Writeplc(tb_Keo2.Text, 0);
                                Thread.Sleep(200);
                            }

                            if (dt.Rows[0].ItemArray[3].ToString() == "2")
                            {
                                PLC_bond.Writeplc(tb_Keo3.Text, 0);
                                Thread.Sleep(200);
                            }
                        }
                    }
                }
                else
                {
                    int tmp = 0;
                    if (dt.Rows.Count > 1)
                    {
                        string[] _code = new string[dt.Rows.Count];
                        string[] status = new string[dt.Rows.Count];
                        string[] time = new string[dt.Rows.Count];
                        string[] Ch = new string[dt.Rows.Count];
                        string maxTime = string.Empty;
                        string CH = string.Empty;

                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            status[i] = dt.Rows[i].ItemArray[0].ToString();
                            _code[i] = dt.Rows[i].ItemArray[1].ToString();
                            time[i] = dt.Rows[i].ItemArray[2].ToString();
                            Ch[i] = dt.Rows[i].ItemArray[3].ToString();
                        }
                        maxTime = time[0];
                        string tg = "";
                        string tg2 = "";
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            for (int j = i + 1; j < dt.Rows.Count; j++)
                            {
                                if (DateTime.Parse(time[i]) < DateTime.Parse(time[j]))
                                {
                                    tg2 = time[i];
                                    time[i] = time[j];
                                    time[j] = tg2;
                                    tg = status[i];
                                    status[i] = status[j];
                                    status[j] = tg;
                                }
                            }                                
                        }

                        if (status[0] == "OK" && status[1] == "OK")
                        {
                            if (PLC_bond.readplc(tb_modelRun.Text) == tb_Cell.Text)
                            {
                                if (dt.Rows[tmp].ItemArray[3].ToString() == "1")
                                {
                                    PLC_bond.Writeplc(tb_Keo1.Text, 1);
                                    Thread.Sleep(200);
                                }
                                if (dt.Rows[tmp].ItemArray[3].ToString() == "2")
                                {
                                    PLC_bond.Writeplc(tb_Keo2.Text, 1);
                                    Thread.Sleep(200);
                                }
                                if (dt.Rows[tmp].ItemArray[3].ToString() == "3")
                                {
                                    PLC_bond.Writeplc(tb_Keo3.Text, 1);
                                    Thread.Sleep(200);
                                }
                            }
                            else if (PLC_bond.readplc(tb_modelRun.Text) == tb_Main.Text)
                            {
                                PLC_bond.Writeplc(tb_Keo1.Text, 0);
                                Thread.Sleep(200);

                                if (dt.Rows[tmp].ItemArray[3].ToString() == "1")
                                {
                                    PLC_bond.Writeplc(tb_Keo2.Text, 1);
                                    Thread.Sleep(200);
                                }
                                if (dt.Rows[tmp].ItemArray[3].ToString() == "2")
                                {
                                    PLC_bond.Writeplc(tb_Keo3.Text, 1);
                                    Thread.Sleep(200);
                                }
                            }
                        }
                        else
                        {
                            if (PLC_bond.readplc(tb_modelRun.Text) == tb_Main.Text)
                            {
                                if (PLC_bond.readplc(tb_DPos.Text) == tb_Pos1.Text)
                                {
                                    PLC_bond.Writeplc(tb_Keo2.Text, 0);
                                }
                                if (PLC_bond.readplc(tb_DPos.Text) == tb_Pos2.Text)
                                {
                                    PLC_bond.Writeplc(tb_Keo3.Text, 0);
                                }
                            }

                            if (PLC_bond.readplc(tb_modelRun.Text) == tb_Cell.Text)
                            {
                                if (PLC_bond.readplc(tb_DPos.Text) == tb_Pos1.Text)
                                {
                                    PLC_bond.Writeplc(tb_Keo1.Text, 0);
                                }
                                if (PLC_bond.readplc(tb_DPos.Text) == tb_Pos2.Text)
                                {
                                    PLC_bond.Writeplc(tb_Keo2.Text, 0);
                                }
                                if (PLC_bond.readplc(tb_DPos.Text) == tb_Pos3.Text)
                                {
                                    PLC_bond.Writeplc(tb_Keo3.Text, 0);
                                }
                            }
                        }
                       
                    }
                    else
                    {
                        if (PLC_bond.readplc(tb_modelRun.Text) == tb_Main.Text)
                        {
                            if (PLC_bond.readplc(tb_DPos.Text) == tb_Pos1.Text)
                            {
                                PLC_bond.Writeplc(tb_Keo2.Text, 0);
                            }
                            if (PLC_bond.readplc(tb_DPos.Text) == tb_Pos2.Text)
                            {
                                PLC_bond.Writeplc(tb_Keo3.Text, 0);
                            }
                        }

                        if (PLC_bond.readplc(tb_modelRun.Text) == tb_Cell.Text)
                        {
                            if (PLC_bond.readplc(tb_DPos.Text) == tb_Pos1.Text)
                            {
                                PLC_bond.Writeplc(tb_Keo1.Text, 0);
                            }
                            if (PLC_bond.readplc(tb_DPos.Text) == tb_Pos2.Text)
                            {
                                PLC_bond.Writeplc(tb_Keo2.Text, 0);
                            }
                            if (PLC_bond.readplc(tb_DPos.Text) == tb_Pos3.Text)
                            {
                                PLC_bond.Writeplc(tb_Keo3.Text, 0);
                            }
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                if (PLC_bond.readplc(tb_modelRun.Text) == tb_Main.Text)
                {
                    if (PLC_bond.readplc(tb_DPos.Text) == tb_Pos1.Text)
                    {
                        PLC_bond.Writeplc(tb_Keo2.Text, 0);
                    }
                    if (PLC_bond.readplc(tb_DPos.Text) == tb_Pos2.Text)
                    {
                        PLC_bond.Writeplc(tb_Keo3.Text, 0);
                    }
                }

                if (PLC_bond.readplc(tb_modelRun.Text) == tb_Cell.Text)
                {
                    if (PLC_bond.readplc(tb_DPos.Text) == tb_Pos1.Text)
                    {
                        PLC_bond.Writeplc(tb_Keo1.Text, 0);
                    }
                    if (PLC_bond.readplc(tb_DPos.Text) == tb_Pos2.Text)
                    {
                        PLC_bond.Writeplc(tb_Keo2.Text, 0);
                    }
                    if (PLC_bond.readplc(tb_DPos.Text) == tb_Pos3.Text)
                    {
                        PLC_bond.Writeplc(tb_Keo3.Text, 0);
                    }
                }

            }
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            string[] array = File.ReadAllLines(@Application.StartupPath + "\\User.txt");
            string[] user = new string[array.Length];
            string[] pass = new string[array.Length];
            bool tt = false;
            for (int i = 0; i < array.Length; i++)
            {
                string[] _array = new string[2];
                _array = array[i].Split('\t');
                user[i] = _array[0];
                pass[i] = _array[1];

            }
            for (int i = 0; i < array.Length; i++)
            {
                if (tb_User.Text == user[i] && tb_Pass.Text == pass[i])
                    tt = true;
            }
            if (bt_Login.Text == "Đăng nhập")
            {
                if (tt == true)
                {
                    tb_Pass.Enabled = false;
                    tb_User.Enabled = false;
                    tb_Pass.Text = "";
                    bt_Login.Text = "Đăng xuất";
                    cb_Model.Enabled = true;
                    button2.Enabled = true;

                    btn_browseMaster.Enabled = true;
                    btn_browseActual.Enabled = true;
                    btn_saveLinkMaster.Enabled = true;

                    txt_fileName.ReadOnly = false;
                    txt_fileNameActual.ReadOnly = false;

                    chb_changeMaster.Enabled = true;

                    cb_Model.Enabled = true;
                    btn_excel.Enabled = true;

                    File.AppendAllText(@Application.StartupPath + "\\Result\\log.txt", "\r\n" + DateTime.Now.ToString() + "\r\nName : " + tb_User.Text + "");
                }
                else
                {
                    MessageBox.Show("Đăng nhập không thành công, bạn hãy đăng nhập lại");
                    tb_User.Text = "";
                    tb_Pass.Text = "";
                }

            }
            else if (bt_Login.Text == "Đăng xuất")
            {
                tb_Pass.Enabled = true;
                tb_User.Enabled = true;
                tb_User.Text = "";
                bt_Login.Text = "Đăng nhập";
                cb_Model.Text = "";
                //button2.Text = "";
                //cb_Model.Enabled = false;
                //button2.Enabled = false;
                timer1.Enabled = false;
                tt = false;
                txt_fileName.ReadOnly = true;
                txt_fileNameActual.ReadOnly = true;

                btn_browseMaster.Enabled = false;
                btn_browseActual.Enabled = false;
                btn_saveLinkMaster.Enabled = false;

                btn_excel.Enabled = false;
                chb_changeMaster.Enabled = false;
                chb_changeMaster.Checked = false;
            }
        }

        int step = 1;
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (PLC_bond.readplc(tb_EndBond.Text) == "1" && PLC_bond.readplc(tb_Enable.Text) == "0")
            {
                Error = 1;
                check_Ready = true;
                tb_Code1.Text = "";
                tb_Code2.Text = "";
                tb_Code3.Text = "";
                vTri1 = true;
                vTri2 = true;
                vTri3 = true;
                PLC_bond.Writeplc(tb_Keo1.Text, 0);
                PLC_bond.Writeplc(tb_Keo2.Text, 0);
                PLC_bond.Writeplc(tb_Keo3.Text, 0);
                step = 1;
            }

            switch (step)
            {
                case 1:
                    if (PLC_bond.readplc(tb_DPos.Text) == tb_Pos1.Text && vTri1 == true && (PLC_bond.readplc(tb_modelRun.Text) == tb_Main.Text || PLC_bond.readplc(tb_modelRun.Text) == tb_Cell.Text))
                    {
                        vTri1 = false;
                        try
                        {
                            if (tb_Code1.Text == "")
                            {
                                PLC_bond.Writeplc(tb_StartRead.Text, 1);
                                Thread.Sleep(200);
                                PLC_bond.Writeplc(tb_StartRead.Text, 0);
                                tb_Code1.Text = Sr1000.Data;
                                Thread.Sleep(200);
                                if (tb_Code1.Text != "" || tb_Code1.Text != "ERROR")
                                {
                                    PLC_bond.Writeplc(tb_EndRead.Text, 1);
                                    Thread.Sleep(200);
                                    PLC_bond.Writeplc(tb_EndRead.Text, 0);
                                    CompareCode(tb_Code1);
                                }
                                else
                                {
                                    PLC_bond.Writeplc(tb_StartRead.Text, 1);
                                    Thread.Sleep(200);
                                    PLC_bond.Writeplc(tb_StartRead.Text, 0);
                                    tb_Code1.Text = Sr1000.Data;
                                    Thread.Sleep(200);
                                    PLC_bond.Writeplc(tb_EndRead.Text, 1);
                                    Thread.Sleep(200);
                                    PLC_bond.Writeplc(tb_EndRead.Text, 0);
                                    CompareCode(tb_Code1);
                                }
                            }
                        }
                        catch (Exception ee)
                        {
                            tb_Code1.Text = "";
                        }
                        step = 2;
                    }
                   
                    break;
                case 2:
                    if (PLC_bond.readplc(tb_DPos.Text) == tb_Pos2.Text && vTri2 == true && (PLC_bond.readplc(tb_modelRun.Text) == tb_Main.Text || PLC_bond.readplc(tb_modelRun.Text) == tb_Cell.Text))
                    {
                        try
                        {
                            vTri2 = false;
                            if (tb_Code2.Text == "")
                            {
                                PLC_bond.Writeplc(tb_StartRead.Text, 1);
                                Thread.Sleep(200);
                                PLC_bond.Writeplc(tb_StartRead.Text, 0);
                                tb_Code2.Text = Sr1000.Data;
                                Thread.Sleep(200);
                                if (tb_Code2.Text != "" || tb_Code2.Text != "ERROR")
                                {
                                    PLC_bond.Writeplc(tb_EndRead.Text, 1);
                                    Thread.Sleep(100);
                                    PLC_bond.Writeplc(tb_EndRead.Text, 0);
                                    CompareCode(tb_Code2);
                                }
                                else
                                {
                                    PLC_bond.Writeplc(tb_StartRead.Text, 1);
                                    Thread.Sleep(300);
                                    PLC_bond.Writeplc(tb_StartRead.Text, 0);
                                    tb_Code2.Text = Sr1000.Data;
                                    Thread.Sleep(200);
                                    PLC_bond.Writeplc(tb_EndRead.Text, 1);
                                    Thread.Sleep(100);
                                    PLC_bond.Writeplc(tb_EndRead.Text, 0);
                                    CompareCode(tb_Code2);
                                }
                                if (PLC_bond.readplc(tb_modelRun.Text) == tb_Main.Text)
                                {
                                    End_check = 1;
                                }
                                else
                                {
                                    // vTri1 = true;
                                }
                            }
                        }
                        catch (Exception ee)
                        {
                            tb_Code2.Text = "";
                        }

                        if (PLC_bond.readplc(tb_modelRun.Text) == tb_Main.Text)
                        {
                            step = 4;
                        }
                        if (PLC_bond.readplc(tb_modelRun.Text) == tb_Cell.Text)
                        {
                            step = 3;
                        }                       
                    }
                    break;
                case 3:
                    if (PLC_bond.readplc(tb_DPos.Text) == tb_Pos3.Text && vTri3 == true && (PLC_bond.readplc(tb_modelRun.Text) == tb_Main.Text || PLC_bond.readplc(tb_modelRun.Text) == tb_Cell.Text))
                    {
                        try
                        {
                            vTri3 = false;
                            if (tb_Code3.Text == "")
                            {
                                PLC_bond.Writeplc(tb_StartRead.Text, 1);
                                Thread.Sleep(200);
                                PLC_bond.Writeplc(tb_StartRead.Text, 0);
                                tb_Code3.Text = Sr1000.Data;
                                Thread.Sleep(200);
                                if (tb_Code3.Text != "" || tb_Code3.Text != "ERROR")
                                {
                                    PLC_bond.Writeplc(tb_EndRead.Text, 1);
                                    Thread.Sleep(100);
                                    PLC_bond.Writeplc(tb_EndRead.Text, 0);
                                    CompareCode(tb_Code3);
                                }
                                else
                                {
                                    PLC_bond.Writeplc(tb_StartRead.Text, 1);
                                    Thread.Sleep(300);
                                    PLC_bond.Writeplc(tb_StartRead.Text, 0);
                                    tb_Code3.Text = Sr1000.Data;
                                    Thread.Sleep(200);
                                    PLC_bond.Writeplc(tb_EndRead.Text, 1);
                                    Thread.Sleep(100);
                                    PLC_bond.Writeplc(tb_EndRead.Text, 0);
                                    CompareCode(tb_Code3);
                                }
                                if (PLC_bond.readplc(tb_modelRun.Text) == tb_Cell.Text)
                                {
                                    End_check = 1;
                                }
                            }
                        }
                        catch (Exception)
                        {
                            tb_Code3.Text = "";
                        }
                        step = 4;
                    }
                    
                    break;
                case 4:
                    if (PLC_bond.readplc(tb_Enable.Text) == "0" && PLC_bond.readplc(tb_EndBond.Text) == "0" && tb_Code1.Text != "" && End_check == 1 && check_Ready == true && (PLC_bond.readplc(tb_modelRun.Text) == tb_Main.Text || PLC_bond.readplc(tb_modelRun.Text) == tb_Cell.Text))
                    {
                        End_check = 0;
                        check_Ready = false;
                        XuLyCode();
                    }                    
                    break;
                case 5:
                    if (PLC_bond.readplc(tb_Enable.Text) == "0" && PLC_bond.readplc(tb_stopbond.Text) == "0" && PLC_bond.readplc(tb_EndBond.Text) == "1" && check_Ready == false && (PLC_bond.readplc(tb_modelRun.Text) == tb_Main.Text || PLC_bond.readplc(tb_modelRun.Text) == tb_Cell.Text))
                    {
                        PLC_bond.Writeplc(tb_Keo1.Text, 0);
                        PLC_bond.Writeplc(tb_Keo2.Text, 0);
                        PLC_bond.Writeplc(tb_Keo3.Text, 0);
                        check_Ready = true;
                        tb_Code1.Text = "";
                        tb_Code2.Text = "";
                        tb_Code3.Text = "";
                        vTri1 = true;
                        vTri2 = true;
                        vTri3 = true;
                        Error = 0;
                        step = 1;
                    }
                    break;
                default:
                    break;
            }

            _memory = PLC.readplc("D1200");
            start = PLC.readplc("M1000");
            Jig_End = PLC.readplc("D1202");
            error = PLC.readplc("D1043");
            
            //----------------------------//            
            End_bond = PLC_bond.readplc(tb_EndBond.Text);
            read_Code1 = PLC_bond.readplc("M190");
            tb_CurPos.Text = PLC_bond.readplc("D50");

            if (DateTime.Now.Second % 5 == 0)
            {
                CH_1 = PLC.readplc("D1013.0");
                CH_2 = PLC.readplc("D1013.1");
                CH_3 = PLC.readplc("D1013.2");
                if (CH_1 == "1" && CH_2 == "1" && CH_3 == "1")
                {
                    CH_Enable = "3";
                }
                else if (CH_1 == "1" && (CH_2 == "1" || CH_3 == "1"))
                {
                    CH_Enable = "2";
                }
                else if (CH_1 == "1" && (CH_2 == "0" && CH_3 == "0"))
                {
                    CH_Enable = "1";
                }
                if (PLC_bond.readplc(tb_Enable.Text) == "0")
                {
                    ts_End.BackColor = Color.Blue;
                }
                else
                {
                    ts_End.BackColor = Color.Red;
                }
            }

            if (start == "1")
            {
                memory = false;
            }           

            if (memory == false && start == "0")
            {
                try
                {
                    memory = true;
                    label10.Text = "OK";
                    modelTest(vtriCode, vtriStatus, vtriTime, vtriCH);
                    output();
                    if (start == "1")
                    {
                        memory = false;
                    }
                }
                catch (Exception)
                {

                }
            }            
        }

        public bool kiemTra(string linkFile1, string linkFile2)
        {
            int columns_special = 0;

            ////khoi tao table1            
            DataTable tbl1 = ReadCsvFile(@linkFile1, 5, 10);
            dgv_master.DataSource = tbl1;

            // khoi tao table2
            DataTable tbl2 = ReadCsvFile(@linkFile2, 5, 10);
            dgv_actual.DataSource = tbl2;

            if (tbl1.Rows.Count != tbl2.Rows.Count || tbl1.Columns.Count != tbl2.Columns.Count)
            {
                return false;
            }

            for (int i = 0; i < tbl1.Rows.Count; i++)
            {
                for (int j = 0; j < tbl1.Columns.Count; j++)
                {
                    if ((tbl1.Rows[i][j].ToString() == "Manufacture Date" && tbl2.Rows[i][j].ToString() == "Manufacture Date"))
                    {
                        columns_special = j;
                    }

                    if (j == columns_special || j == 90 || j == PCM_Date || tbl1.Rows[i][j].ToString() == "\n\t" || tbl1.Rows[i][j].ToString() == "" || tbl2.Rows[i][j].ToString() == "\n\t" || tbl2.Rows[i][j].ToString() == "")
                    {
                        string a = tbl1.Rows[i][j].ToString();
                        string b = tbl2.Rows[i][j].ToString();
                    }
                    else
                    {
                        if (!Equals(tbl1.Rows[i][j], tbl2.Rows[i][j]))
                        {
                            string a = tbl1.Rows[i][j].ToString();
                            string b = tbl2.Rows[i][j].ToString();
                            dgv_master.Rows[i].Cells[j].Style.BackColor = Color.Red;
                            dgv_actual.Rows[i].Cells[j].Style.BackColor = Color.Red;
                            return false;
                        }
                    }
                }
            }
            return true;
        }

        public bool TimeBetween(DateTime time, DateTime startDateTime, DateTime endDateTime)
        {
            // get TimeSpan
            TimeSpan start = new TimeSpan(startDateTime.Hour, startDateTime.Minute, 0);
            TimeSpan end = new TimeSpan(endDateTime.Hour, endDateTime.Minute, 0);

            // convert datetime to a TimeSpan
            TimeSpan now = time.TimeOfDay;
            // see if start comes before end
            if (start < end)
                return start <= now && now <= end;
            // start is after end, so do the inverse comparison
            return !(end < now && now < start);
        }

        public string find_shift()
        {
            string shift;
            DateTime dateTime = DateTime.Now;

            DateTime startDateTime = new DateTime(dateTime.Year, dateTime.Month, dateTime.Day, 8, 0, 0);
            DateTime endDateTime = new DateTime(dateTime.Year, dateTime.Month, dateTime.Day, 20, 0, 0);

            if (TimeBetween(dateTime, startDateTime, endDateTime))
            {
                shift = "Day";
            }
            else
            {
                shift = "Night";
            }
            return shift;
        }

        private void btn_excel_Click(object sender, EventArgs e)
        {
            string fileName = Application.StartupPath + "\\Result\\"
                + DateTime.Now.Month + "-" + DateTime.Now.Day + "-" + DateTime.Now.Year + "_" + find_shift() + "_" + cb_Model.Text + "";

            ExportToExcel(dgv_master, dgv_actual, fileName);
        }

        public void ExportToExcel(DataGridView dgv_master, DataGridView dgv_actual, string fileName)
        {
            try
            {
                if (dgv_master == null || dgv_master.Rows.Count == 0 || dgv_actual == null || dgv_actual.Rows.Count == 0)
                {
                    throw new Exception("ExportToExcel: Null or empty input table!\n");
                }

                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
                app.Visible = false;
                worksheet = workbook.Sheets["Sheet1"];
                worksheet = workbook.ActiveSheet;
                worksheet.Name = "Records";

                try
                {
                    for (int headerId = 1; headerId < dgv_master.Columns.Count + 1; headerId++)
                    {
                        if (dgv_master.Columns[headerId - 1].HeaderText.Contains("Column"))
                        {
                            worksheet.Cells[6, headerId] = String.Empty;
                        }
                        else
                        {
                            worksheet.Cells[6, headerId] = dgv_master.Columns[headerId - 1].HeaderText;
                        }

                    }
                    int i = 0;
                    int j = 0;
                    for (i = 0; i < dgv_master.Rows.Count - 0; i++)
                    {
                        for (j = 0; j < dgv_master.Columns.Count; j++)
                        {
                            if (dgv_master.Rows[i].Cells[j].Value != null)
                            {
                                worksheet.Cells[i + 7, j + 1] = dgv_master.Rows[i].Cells[j].Value.ToString();
                            }
                            else
                            {
                                worksheet.Cells[i + 7, j + 1] = "";
                            }
                        }
                    }

                    worksheet.Cells[1, 1] = DateTime.Now.ToString();
                    worksheet.Cells[2, 1] = "Nguoi kiem tra : " + tb_User.Text;
                    worksheet.Cells[3, 1] = "Model : " + cb_Model.Text;
                    worksheet.Cells[4, 1] = result.ToString();
                    worksheet.Cells[5, 1] = "Check sheet master";
                    worksheet.Cells[12, 1] = "Check sheet actual";

                    for (int headerId = 1; headerId < dgv_actual.Columns.Count + 1; headerId++)
                    {
                        if (dgv_actual.Columns[headerId - 1].HeaderText.Contains("Column"))
                        {
                            worksheet.Cells[13, headerId] = String.Empty;
                        }
                        else
                        {
                            worksheet.Cells[13, headerId] = dgv_actual.Columns[headerId - 1].HeaderText;
                        }
                    }
                    for (int m = 0; m < dgv_actual.Rows.Count - 0; m++)
                    {
                        for (int n = 0; n < dgv_actual.Columns.Count; n++)
                        {
                            if (dgv_actual.Rows[m].Cells[n].Value != null)
                            {
                                worksheet.Cells[m + 14, n + 1] = dgv_actual.Rows[m].Cells[n].Value.ToString();
                            }
                            else
                            {
                                worksheet.Cells[m + 14, n + 1] = "";
                            }
                        }
                    }


                    //Excel.Range tRange = worksheet[1,1].UsedRange;
                    //tRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    //tRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

                    //tRange.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    //tRange.Columns.AutoFit();

                    //SaveFileDialog saveDialog = new SaveFileDialog();
                    //saveDialog.Filter = "CSV files (*.csv)|*.csv|Excel |*.xlsx|Excel 2003 | *.xls";
                    //saveDialog.FilterIndex = 2;

                    //if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    //{
                    //    workbook.SaveAs(saveDialog.FileName);
                    //    MessageBox.Show("Save excel Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //}

                    if (!string.IsNullOrEmpty(fileName + ".csv"))
                    {
                        try
                        {
                            worksheet.SaveAs(fileName + ".csv");
                            app.Quit();
                            MessageBox.Show("Đã lưu thành công tại " + Application.StartupPath + "\\Result\\"
                + DateTime.Now.Month + "-" + DateTime.Now.Day + "-" + DateTime.Now.Year + " " + DateTime.Now.Hour + "-"
                + DateTime.Now.Minute + "-" + DateTime.Now.Second + "_" + cb_Model.Text + "_" + tb_User.Text + "");
                        }
                        catch (Exception ex)
                        {
                            throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"
                                            + ex.Message);
                        }
                    }
                    else
                    {
                        // no file path is given
                        app.Visible = true;
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                finally
                {
                    app.Quit();
                    workbook = null;
                    worksheet = null;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btn_importExcel_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();

            if (result == DialogResult.OK)
            {

                string fileBrowse = openFileDialog1.FileName;
                txt_fileName.Text = fileBrowse;

                try
                {
                    if (txt_fileNameActual.Text == "")
                    {
                        dgv_master.Columns.Clear();
                        dgv_actual.Columns.Clear();
                        DataTable tbl1 = ReadCsvFile(@txt_fileName.Text, 5, 10);
                        dgv_master.DataSource = tbl1;
                    }
                    else
                    {
                        dgv_master.Columns.Clear();
                        dgv_actual.Columns.Clear();
                        kq = kiemTra(txt_fileName.Text, txt_fileNameActual.Text);

                        if (kq == true)
                        {
                            lb_pass.BackColor = Color.Green;
                            lb_pass.Text = "OK";
                            MessageBox.Show("Master thông số OK");

                        }
                        else
                        {
                            lb_pass.BackColor = Color.Red;
                            lb_pass.Text = "NG";
                            MessageBox.Show("Master thông số NG. Hãy kiểm tra lại!", "Chú ý", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                catch (Exception)
                {

                }
            }
        }

        private void btn_browseActual_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog2.ShowDialog();

            if (result == DialogResult.OK)
            {
                string fileBrowse = openFileDialog2.FileName;
                txt_fileNameActual.Text = fileBrowse;

                try
                {
                    if (txt_fileName.Text == "")
                    {
                        dgv_master.Columns.Clear();
                        dgv_actual.Columns.Clear();
                        DataTable tbl1 = ReadCsvFile(@txt_fileNameActual.Text, 5, 10);
                        dgv_actual.DataSource = tbl1;
                    }
                    else
                    {
                        dgv_actual.Columns.Clear();
                        dgv_master.Columns.Clear();
                        kq = kiemTra(txt_fileName.Text, txt_fileNameActual.Text);

                        if (kq == true)
                        {
                            lb_pass.BackColor = Color.Green;
                            lb_pass.Text = "OK";
                            MessageBox.Show("Master thông số OK");
                        }
                        else
                        {
                            lb_pass.BackColor = Color.Red;
                            lb_pass.Text = "NG";
                            MessageBox.Show("Master thông số NG. Hãy kiểm tra lại!", "Chú ý", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                catch (Exception)
                {

                }
            }
        }

        private void btn_saveLinkMaster_Click(object sender, EventArgs e)
        {
            try
            {
                if (chb_changeMaster.Checked == true)
                {
                    if (cb_Model.Text == "" || txt_fileName.Text == "")
                    {
                        MessageBox.Show("Hãy chọn model trước khi lưu link master hoặc link file name đang trống");
                    }
                    else if (cb_Model.Text == "P01P-00226A")
                    {
                        DialogResult traloi;
                        traloi = MessageBox.Show("Bạn có chắc chắn lưu link master thông số " + cb_Model.Text + " này không??", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        File.WriteAllText("LinkMaster_P01P-00226A.txt", String.Empty);
                        File.AppendAllText("LinkMaster_P01P-00226A.txt", txt_fileName.Text);

                        MessageBox.Show("Lưu thành công!");

                        File.AppendAllText(@Application.StartupPath + "\\Result\\log.txt", "\r\n"
                                + DateTime.Now.ToString() + "\r\nName : " + tb_User.Text + "\r Thay đổi master " + cb_Model.Text + " ");
                        //txt_fileName.Text = "";
                        //txt_fileNameActual.Text = "";
                    }
                    else if (cb_Model.Text == "P01P-00225A")
                    {
                        DialogResult traloi;
                        traloi = MessageBox.Show("Bạn có chắc chắn lưu link master thông số " + cb_Model.Text + " này không??", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        File.WriteAllText("LinkMaster_P01P-00225A.txt", String.Empty);
                        File.AppendAllText("LinkMaster_P01P-00225A.txt", txt_fileName.Text);
                        MessageBox.Show("Lưu thành công!");

                        File.AppendAllText(@Application.StartupPath + "\\Result\\log.txt", "\r\n"
                                + DateTime.Now.ToString() + "\r\nName : " + tb_User.Text + "\r Thay đổi master " + cb_Model.Text + " ");

                        //txt_fileName.Text = "";
                        //txt_fileNameActual.Text = "";
                    }
                    else if (cb_Model.Text == "P01P-00227A")
                    {
                        DialogResult traloi;
                        traloi = MessageBox.Show("Bạn có chắc chắn lưu link master thông số " + cb_Model.Text + " này không??", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        File.WriteAllText("LinkMaster_P01P-00227A.txt", String.Empty);
                        File.AppendAllText("LinkMaster_P01P-00227A.txt", txt_fileName.Text);
                        MessageBox.Show("Lưu thành công!");

                        File.AppendAllText(@Application.StartupPath + "\\Result\\log.txt", "\r\n"
                                + DateTime.Now.ToString() + "\r\nName : " + tb_User.Text + "\r Change master " + cb_Model.Text + " ");


                    }
                    else
                    {

                    }
                }
                else
                {
                    MessageBox.Show("Để thay đổi master cần click chọn check box Đổi Master -> chọn Model cần đổi");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgv_master_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewColumn column in dgv_master.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }
        private void dgv_actual_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewColumn column in dgv_actual.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }
        public DataTable Table_new(string[] result, string[] barcode, string[] time, string[] CH)
        {
            DataTable TableExcel = new DataTable();
            DataColumn column;
            DataRow Row;

            column = new DataColumn();
            column.DataType = typeof(String);
            column.ColumnName = "Result";
            TableExcel.Columns.Add(column);

            column = new DataColumn();
            column.DataType = typeof(String);
            column.ColumnName = "Barcode";
            TableExcel.Columns.Add(column);

            column = new DataColumn();
            column.DataType = typeof(String);
            column.ColumnName = "Time";
            TableExcel.Columns.Add(column);

            column = new DataColumn();
            column.DataType = typeof(String);
            column.ColumnName = "CH";
            TableExcel.Columns.Add(column);


            if (result.Length > 0 && barcode.Length > 0 && CH.Length > 0)
            {
                for (int i = 0; i < result.Length; i++)
                {
                    Row = TableExcel.NewRow();
                    Row["Result"] = result[i];
                    Row["Barcode"] = barcode[i];
                    Row["Time"] = time[i];
                    Row["CH"] = CH[i];
                    TableExcel.Rows.Add(Row);
                }
            }

            if (result.Length == 0)
            {
                Row = TableExcel.NewRow();
                Row["Result"] = "";
                Row["Barcode"] = "";
                Row["Time"] = "";
                Row["CH"] = "";
                TableExcel.Rows.Add(Row);
            }
            return TableExcel;
        }
        private void SR1000_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            codeSR += SR1000.ReadExisting();
        }
        private DateTime dateTime_Actual()
        {
            DateTime dateTime = DateTime.Now;
            DateTime startDateTime = new DateTime(dateTime.Year, dateTime.Month, dateTime.Day, 7, 45, 0);
            DateTime endDateTime = new DateTime(dateTime.Year, dateTime.Month, dateTime.Day, 19, 44, 0);

            DateTime dateTimeNight = DateTime.Now.AddDays(1);
            DateTime startDateTimeNight = new DateTime(dateTime.Year, dateTime.Month, dateTime.Day, 19, 45, 0);
            DateTime endDateTimeNight = new DateTime(dateTimeNight.Year, dateTimeNight.Month, dateTimeNight.Day, 7, 44, 0);
            if (dateTime > startDateTime && dateTime < endDateTime)
            {
                return dateTime;
            }
            else
            {
                return dateTimeNight;
            }

            //if (dateTime > startDateTimeNight && dateTime < endDateTimeNight)
            //{
            //    return dateTimeNight;
            //}
        }
        private void Create_File()
        {
            DateTime dateTime = DateTime.Now;
            DateTime startDateTime = new DateTime(dateTime.Year, dateTime.Month, dateTime.Day, 7, 45, 0);
            DateTime endDateTime = new DateTime(dateTime.Year, dateTime.Month, dateTime.Day, 19, 44, 0);

            DateTime dateTimeNight = DateTime.Now.AddDays(1);
            DateTime startDateTimeNight = new DateTime(dateTime.Year, dateTime.Month, dateTime.Day, 19, 45, 0);
            DateTime endDateTimeNight = new DateTime(dateTimeNight.Year, dateTimeNight.Month, dateTimeNight.Day, 7, 44, 0);
            if (dateTime > startDateTime && dateTime < endDateTime)
            {
                if (!File.Exists(@Application.StartupPath + "\\Log\\" + DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString("00") + "\\" + DateTime.Now.Month.ToString("00") + "-" + DateTime.Now.Day.ToString("00") + cb_Model.Text + ".csv"))
                {
                    Excel.Export_CSV(Table_new(OK_NG, barcode, time_Test, CH), @Application.StartupPath + "\\Log\\" + DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString("00") + "\\" + "[Day-" + cb_Model.Text + "]" + DateTime.Now.Month.ToString("00") + "-" + DateTime.Now.Day.ToString("00") + ".csv", false, "Result, Barcode, CH\n");
                }
            }


            if (dateTime > startDateTimeNight && dateTime < endDateTimeNight)
            {
                if (!File.Exists(@Application.StartupPath + "\\Log\\" + dateTimeNight.Year.ToString() + "-" + dateTimeNight.Month.ToString("00") + "\\" + dateTimeNight.Month.ToString("00") + "-" + dateTimeNight.Day.ToString("00") + cb_Model.Text + ".csv"))
                {
                    Excel.Export_CSV(Table_new(OK_NG, barcode, time_Test, CH), @Application.StartupPath + "\\Log\\" + dateTimeNight.Year.ToString() + "-" + dateTimeNight.Month.ToString("00") + "\\" + "[Night-" + cb_Model.Text + "]" + dateTimeNight.Month.ToString("00") + "-" + dateTimeNight.Day.ToString("00") + ".csv", false, "Result, Barcode, CH\n");
                }
            }
        }

        private void btn_change_Click(object sender, EventArgs e)
        {
            string[] arry = File.ReadAllLines(@Application.StartupPath + "\\Config.txt");
            for (int i = 0; i < arry.Length; i++)
            {
                if (arry[i].Contains(lb_StartRead.Text))
                {
                    arry[i] = lb_StartRead.Text + "|" + tb_StartRead.Text;
                }

                if (arry[i].Contains(lb_EndRead.Text))
                {
                    arry[i] = lb_EndRead.Text + "|" + tb_EndRead.Text;
                }

                if (arry[i].Contains(lb_pos1.Text))
                {
                    arry[i] = lb_pos1.Text + "|" + tb_Pos1.Text;
                }

                if (arry[i].Contains(lb_pos2.Text))
                {
                    arry[i] = lb_pos2.Text + "|" + tb_Pos2.Text;
                }

                if (arry[i].Contains(lb_pos3.Text))
                {
                    arry[i] = lb_pos3.Text + "|" + tb_Pos3.Text;
                }

                if (arry[i].Contains(lb_DPos.Text))
                {
                    arry[i] = lb_DPos.Text + "|" + tb_DPos.Text;
                }

            }
            File.WriteAllLines(@Application.StartupPath + "\\Config.txt", arry);
        }
        private void btn_Save_Click(object sender, EventArgs e)
        {
            string[] arry = File.ReadAllLines(@Application.StartupPath + "\\Config.txt");
            for (int i = 0; i < arry.Length; i++)
            {
                if (arry[i].Contains(lb_StartBond.Text) || arry[i].Contains(lb_EndBond.Text) || arry[i].Contains(lb_modelRun.Text) || arry[i].Contains(lb_stopbond.Text))
                {
                    if (arry[i].Contains(lb_StartBond.Text))
                    {
                        arry[i] = lb_StartBond.Text + "|" + tb_Startbond.Text;
                    }
                    if (arry[i].Contains(lb_linkdata.Text))
                    {
                        arry[i] = lb_linkdata.Text + "|" + tb_Enable.Text;
                    }
                    if (arry[i].Contains(lb_EndBond.Text))
                    {
                        arry[i] = lb_EndBond.Text + "|" + tb_EndBond.Text;
                    }

                    if (arry[i].Contains(lb_modelRun.Text))
                    {
                        arry[i] = lb_modelRun.Text + "|" + tb_modelRun.Text;
                    }

                    if (arry[i].Contains(lb_stopbond.Text))
                    {
                        arry[i] = lb_stopbond.Text + "|" + tb_stopbond.Text;
                    }
                }
            }
            File.WriteAllLines(@Application.StartupPath + "\\Config.txt", arry);
        }
        private void btn_save2_Click(object sender, EventArgs e)
        {
            string[] arry = File.ReadAllLines(@Application.StartupPath + "\\Config.txt");
            for (int i = 0; i < arry.Length; i++)
            {
                if (arry[i].Contains(lb_Keo1.Text) || arry[i].Contains(lb_Keo2.Text) || arry[i].Contains(lb_Keo3.Text))
                {
                    if (arry[i].Contains(lb_Keo1.Text))
                    {
                        arry[i] = lb_Keo1.Text + "|" + tb_Keo1.Text;
                    }

                    if (arry[i].Contains(lb_Keo2.Text))
                    {
                        arry[i] = lb_Keo2.Text + "|" + tb_Keo2.Text;
                    }

                    if (arry[i].Contains(lb_Keo3.Text))
                    {
                        arry[i] = lb_Keo3.Text + "|" + tb_Keo3.Text;
                    }
                }
            }
            File.WriteAllLines(@Application.StartupPath + "\\Config.txt", arry);
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            modelTest(vtriCode, vtriStatus, vtriTime, vtriCH);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            XuLyCode();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string[] arry = File.ReadAllLines(@Application.StartupPath + "\\Config.txt");
            for (int i = 0; i < arry.Length; i++)
            {

                if (arry[i].Contains(lb_Main.Text))
                {
                    arry[i] = lb_Main.Text + "|" + tb_Main.Text;
                }

                if (arry[i].Contains(lb_Cell.Text))
                {
                    arry[i] = lb_Cell.Text + "|" + tb_Cell.Text;
                }

            }
            File.WriteAllLines(@Application.StartupPath + "\\Config.txt", arry);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            tb_Code1.Text = "";
            tb_Code2.Text = "";
            tb_Code3.Text = "";
            vTri1 = true;
            read_position = 0;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string[] arry = File.ReadAllLines(@Application.StartupPath + "\\Config.txt");
            for (int i = 0; i < arry.Length; i++)
            {

                if (arry[i].Contains(lb_PLCFuntion.Text))
                {
                    arry[i] = lb_PLCFuntion.Text + "|" + tb_IP1.Text;
                }

                if (arry[i].Contains(lb_PLCBonding.Text))
                {
                    arry[i] = lb_PLCBonding.Text + "|" + tb_IP2.Text;
                }

                if (arry[i].Contains(lb_ComScanner.Text))
                {
                    arry[i] = lb_ComScanner.Text + "|" + tb_Scanner.Text;
                }

            }
            File.WriteAllLines(@Application.StartupPath + "\\Config.txt", arry);
        }

        private void bt_CntPlcFct_Click(object sender, EventArgs e)
        {
            PLC.thietlap(tb_IP1);
            PLC.ketnoi(lb_PLC, bt_CntPlcFct);
        }

        private void bt_CntPlcBon_Click(object sender, EventArgs e)
        {
            PLC_bond.thietlap(tb_IP2);
            PLC_bond.ketnoi(lb_PLCbond, bt_CntPlcBon);
        }

        private void bt_CntScan_Click(object sender, EventArgs e)
        {
            Sr1000 = new clsScanner(this);
            Sr1000.COMnum = tb_Scanner.Text;
            Sr1000.ketnoi(bt_CntScan);
        }

        private void tb_Repair_Click(object sender, EventArgs e)
        {
            DialogResult result;
            if (tb_Repair.Text == "Enable")
            {
                result = MessageBox.Show("Bạn có muốn thay đổi thông số", "Infor", MessageBoxButtons.YesNo);
                if (result == System.Windows.Forms.DialogResult.Yes)
                {
                    tb_Enable.Enabled = true;
                    btn_save2.Enabled = true;
                    tb_Repair.Text = "Disable";
                    tb_IP1.Enabled = true;
                    tb_IP2.Enabled = true;
                    tb_Scanner.Enabled = true;
                    tb_StartRead.Enabled = true;
                    tb_EndRead.Enabled = true;
                    tb_DPos.Enabled = true;
                    tb_Pos1.Enabled = true;
                    tb_Pos2.Enabled = true;
                    tb_Pos3.Enabled = true;
                    tb_modelRun.Enabled = true;
                    tb_StartRead.Enabled = true;
                    tb_EndBond.Enabled = true;
                    tb_stopbond.Enabled = true;
                    //tb_CurPos.Enabled = true;
                    tb_Main.Enabled = true;
                    tb_Cell.Enabled = true;
                    tb_Keo1.Enabled = true;
                    tb_Keo2.Enabled = true;
                    tb_Keo3.Enabled = true;
                    button6.Enabled = true;
                    button4.Enabled = true;
                    btn_Save.Enabled = true;
                    btn_change.Enabled = true;
                    tb_Startbond.Enabled = true;
                    //tb_Reset.Enabled = true;
                }
            }
            else
            {
                tb_Enable.Enabled = false;
                //tb_Reset.Enabled = false;
                tb_Repair.Text = "Enable";
                tb_IP1.Enabled = false;
                tb_IP2.Enabled = false;
                tb_Scanner.Enabled = false;
                tb_StartRead.Enabled = false;
                tb_EndRead.Enabled = false;
                tb_DPos.Enabled = false;
                tb_Pos1.Enabled = false;
                tb_Pos2.Enabled = false;
                tb_Pos3.Enabled = false;
                tb_modelRun.Enabled = false;
                tb_StartRead.Enabled = false;
                tb_EndBond.Enabled = false;
                tb_stopbond.Enabled = false;
                //tb_CurPos.Enabled = false;
                tb_Main.Enabled = false;
                tb_Cell.Enabled = false;
                tb_Keo1.Enabled = false;
                tb_Keo2.Enabled = false;
                tb_Keo3.Enabled = false;
                button6.Enabled = false;
                button4.Enabled = false;
                btn_Save.Enabled = false;
                btn_change.Enabled = false;
                tb_Startbond.Enabled = false;
                btn_save2.Enabled = false;
            }
        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            CompareCode(tb_Code1);
        }

        private void button1_Click_3(object sender, EventArgs e)
        {
            checkcod.SaveList("1", "CELL.log", link_Server);
        }

        private void button1_Click_4(object sender, EventArgs e)
        {
            modelTest(vtriCode, vtriStatus, vtriTime, vtriCH);
        }

        private void button1_Click_5(object sender, EventArgs e)
        {
           
        }

        private void button1_Click_6(object sender, EventArgs e)
        {
             
        }

        private void button1_Click_7(object sender, EventArgs e)
        {
            kiemTra(@"D:\Test\[CELK]2022-07-08 12-44-50 (1).csv", @"D:\Test\fdsafasdfssadfsdf.csv");
        }

        public int countTime = 0;
        private void timer2_Tick(object sender, EventArgs e)
        {
            countTime++;
            if (countTime == 300)
            {
                CompareLog();
                countTime = 0;
            }
        }

        public void CompareLog()
        {
            //Boxing logfile
            string[] strSource = GetFile();
            string[] arrSource = new string[1000];
            int count_Source = 0;
            for (int i = 0; i < strSource.Length; i++)
            {
                if (strSource[i] == null)
                {
                    break;
                }

                if (File.Exists(strSource[i]) == true)
                {
                    StreamReader srSource = new StreamReader(strSource[i]);
                    while (srSource.EndOfStream == false)
                    {
                        string str1 = srSource.ReadLine();
                        string[] arrStr1 = str1.Split(',');
                        arrSource[count_Source] = arrStr1[3];
                        count_Source++;
                    }
                    srSource.Close();
                }
            }

            //Input logfile           
            string strDestination = GetPath();
            string[] arrDestination = new string[1000];
            if (File.Exists(strDestination) == true)
            {
                int count_Destination = 0;
                StreamReader srDestination = new StreamReader(strDestination);
                while (srDestination.EndOfStream == false)
                {
                    string str2 = srDestination.ReadLine();
                    string[] arrStr2 = str2.Split('|');
                    string z = arrStr2[1].Substring(4, 2) + "/" + arrStr2[1].Substring(6, 2) + "/" + arrStr2[1].Substring(0, 4) + " " + arrStr2[1].Substring(8, 2) + ":" + arrStr2[1].Substring(10, 2) + ":" + arrStr2[1].Substring(12, 2);
                    int x = DateTime.Compare(Convert.ToDateTime(z), DateTime.Now.AddHours(-6));
                    if (DateTime.Compare(Convert.ToDateTime(z), DateTime.Now.AddHours(-6)) < 0)//input < boxing - 6h
                    {
                        arrDestination[count_Destination] = arrStr2[0];
                        count_Destination++;
                    }
                }
                srDestination.Close();
            }

            //So sánh
            if ((arrSource[0] != null) && (arrDestination[0] != null))
            {
                string[] arrAlarm = new string[1000];
                int count_Alarm = 0;
                for (int i = 0; i < arrDestination.Length; i++)
                {
                    int same = 0;
                    for (int j = 0; j < arrSource.Length; j++)
                    {
                        if (arrDestination[i] == arrSource[j])//input logfile có trong boxing logfile
                        {
                            same++;
                        }

                        if (arrSource[j] == null)
                        {
                            break;
                        }
                    }

                    if (same == 0)//input ko có trong boxing
                    {
                        arrAlarm[count_Alarm] = arrDestination[i];
                        count_Alarm++;
                    }

                    if (arrDestination[i] == null)
                    {
                        break;
                    }
                }

                if (arrAlarm[0] != null)
                {
                    //string toDisply = string.Join("\n", arrAlarm);
                    string toDisply = string.Empty;
                    for (int i = 0; i < arrAlarm.Length; i++)
                    {
                        if (arrAlarm[i] != null)
                        {
                            if (i == 0)
                            {
                                toDisply += arrAlarm[i];
                            }
                            else
                            {
                                toDisply += "\n" + arrAlarm[i];
                            }
                        }

                    }
                    MessageBox.Show("Barcode PCM không FI-FO :\n" + toDisply, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        public string GetPath()
        {
            string str = string.Empty;
            string subStr = string.Empty;
            StreamReader sr = new StreamReader(@Application.StartupPath + "\\LinkInput.txt");
            while (sr.EndOfStream == false)
            {
                subStr = sr.ReadLine();
            }
            sr.Close();
            string dt = DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString() + "-" + DateTime.Now.Day.ToString();
            string molNam = string.Empty;
            if (model_Run == "MAIN")
            {
                molNam = "Composite Module_main";
            }
            else if (model_Run == "CELL")
            {
                molNam = "Cell Module 14S_21700";
            }
            else if (model_Run == "SUB")
            {
                molNam = "Composite Module_sub";
            }
            str = subStr + "\\" + molNam + "\\" + dt + "-OnMes.log";
            return str;
        }

        public string[] GetFile()
        {
            string[] rel = new string[100];
            string toDay = DateTime.Now.ToString("yyyy-MM-dd");
            string yesterDay = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
            string link = System.IO.File.ReadAllText("Link.txt");
            int count_file = 0;
            if (toDay.Substring(8, 2) == "01")
            {
                string[] folder = new string[2]
                    {link + "\\" + DateTime.Now.ToString("yyyy-MM"),
                     link + "\\" + DateTime.Now.AddMonths(-1).ToString("yyyy-MM")
                    };
                for(int i = 0; i < folder.Length; i++)
                {
                    DirectoryInfo dir = new DirectoryInfo(folder[i]);
                    if (dir.Exists)
                    {
                        foreach (FileInfo fIn in dir.GetFiles())
                        {                            
                            if (fIn.Name.Contains(toDay) || fIn.Name.Contains(yesterDay))
                            {
                                rel[count_file] = link + "\\" + DateTime.Now.ToString("yyyy-MM") + "\\" + fIn.Name;
                                count_file++;
                            }
                        }
                    }    
                }
            }
            else
            {
                DirectoryInfo dir = new DirectoryInfo(link + "\\" + DateTime.Now.ToString("yyyy-MM"));
                if (dir.Exists)
                {
                    foreach (FileInfo fIn in dir.GetFiles())
                    {
                        if (fIn.Name.Contains(toDay) || fIn.Name.Contains(yesterDay))
                        {
                            rel[count_file] = link + "\\" + DateTime.Now.ToString("yyyy-MM") + "\\" + fIn.Name;
                            count_file++;
                        }
                    }
                }    
            }
            return rel;
        }   
    }

}



