using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Text;
using System.Drawing;
using System.Windows.Forms;

namespace Auto_Lock
{
    public class clsScanner
    {
        public event SerialDataReceivedEventHandler Datareceived;
        SerialPort Scanner;
        Form1 _frm;
        //clsdataconvert dataconvert;

        private string _data;

        public string Data
        {
            get { return _data; }
            set { _data = value; }
        }

        private string _COMnum;

        public string COMnum
        {
            get { return _COMnum; }
            set { _COMnum = value; }
        }

        public clsScanner(Form1 frm)
        {
            _frm = frm;
            Scanner = new SerialPort();
            //dataconvert = new clsdataconvert();
        }

        public bool ketnoi(Button bt)
        {
            try
            {
                Scanner.PortName = _COMnum;
                Scanner.BaudRate = 9600;
                Scanner.DataBits = 8;
                Scanner.ReadBufferSize = 1024;
                Scanner.WriteBufferSize = 512;
                Scanner.Parity = Parity.None;
                Scanner.DtrEnable = true;
                Scanner.DataReceived += Scanner_DataReceived;
                Scanner.Open();
                bt.BackColor = Color.Green;
                return true;
                    
                           
            }
            catch (Exception)
            {
                Scanner.Close();
                bt.BackColor = Color.Red;
                return false;
            }
            
        }

        public void ngatketnoi()
        {
            try
            {
                Scanner.Close();
            }
            catch (Exception)
            {
                
            }
        }

        private void Scanner_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                _data = Scanner.ReadLine();
                if (Datareceived != null)
                {
                    Datareceived(this, e);
                }
            }
            catch (Exception)
            {                
                //throw;
            }           
        }
    }
}
