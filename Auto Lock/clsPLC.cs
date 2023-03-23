using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ACTETHERLib;
using System.Threading;


namespace Auto_Lock
{
    public class clsPLC
    {
        //public ACTETHERLib.ActFXENETTCP PLC = new ACTETHERLib.ActFXENETTCP();
        //public ActMLQNUDECPUTCP PLC = new ActMLQNUDECPUTCP();
        public ACTETHERLib.ActQNUDECPUTCP PLC = new ACTETHERLib.ActQNUDECPUTCP();
        public ACTETHERLib.ActQJ71E71TCP PLC_bond = new ACTETHERLib.ActQJ71E71TCP();
        private int IRet = 0;
        private int IRet_1 = 0;
        private bool _PLC_flag = false;    

        private int _ActCpuType1;
        private int _ActDestinationPortNumber1;
        private string _ActHostAddress1;
        private int _ActTimeOut1;
        public int ActCpuType1
        {
            get { return _ActCpuType1; }
            set { _ActCpuType1 = value; }
        }
        public int ActDestinationPortNumber1
        {
            get { return _ActDestinationPortNumber1; }
            set { _ActDestinationPortNumber1 = value; }
        }
        public string ActHostAddress1
        {
            get { return _ActHostAddress1; }
            set { _ActHostAddress1 = value; }
        }
        public int ActTimeOut1
        {
            get { return _ActTimeOut1; }
            set { _ActTimeOut1 = value; }
        }
        public clsPLC()
        {
           // thietlap();
        }
        public bool PLC_flag
        {
            get { return _PLC_flag; }
            set { _PLC_flag = value; }
        }

        public string getData(string add)
        {
            int result;
            string data = string.Empty;
            PLC.GetDevice(add, out result);
            data = result.ToString();
            return data;
        }
        public string readplc(string address)
        {
            string adrall = address;
            string[] adr = adrall.Split('\n');
            int IRET_read;
            int[] addlength = new int[adr.Length];
            
            IRET_read = PLC.ReadDeviceRandom(adrall, adr.Length, out addlength[0]);
           
            if (IRET_read == 0)
            {
                return addlength[0].ToString();
            }
            else
            {
                return "FAIL";
            }
        }
        
        public void Writeplc(string address, int value)
        {
            string adrall = address;
            string[] adr = adrall.Split('\n');
            int IRET_read;
            int[] addlength = new int[adr.Length];
            addlength[0] = value;
            IRET_read = PLC.WriteDeviceRandom(adrall, adr.Length, ref addlength[0]);
        }
        //public bool ketnoi(Label lbPLCstatus)
        //{
        //    IRet = PLC.Open();
        //    if (IRet == 0)
        //    {
        //        //lbPLCstatus.Text = "CONNECTED";
        //        lbPLCstatus.BackColor = Color.Green;
        //        _PLC_flag = true;
        //        return true;
        //    }
        //    else
        //    {
        //        //lbPLCstatus.Text = "DISCONNECTED";
        //        lbPLCstatus.BackColor = Color.Red;
        //        _PLC_flag = false;
        //        return false;
        //    }
        //}
        public bool ketnoi(ToolStripStatusLabel lbPLCstatus, Button bt)
        {
            IRet = PLC.Open();
            if (IRet == 0)
            {
                //lbPLCstatus.Text = "CONNECTED";
                lbPLCstatus.BackColor = Color.Blue;
                bt.BackColor = Color.Blue;
                _PLC_flag = true;
                return true;
            }
            else
            {
                //lbPLCstatus.Text = "DISCONNECTED";
                lbPLCstatus.BackColor = Color.Red;
                bt.BackColor = Color.Red;
                _PLC_flag = false;
                return false;
            }
        }

    
        
        public void thietlap(TextBox IP_PLC)
        {
            PLC.ActUnitNumber = 26;
            PLC.ActNetworkNumber = 1;
            PLC.ActStationNumber = 1;            
            PLC.ActIONumber = 1023;
            PLC.ActCpuType = 144;
            //PLC.ActSourceNetworkNumber = 2;
            //PLC.ActSourceStationNumber = 1;
            PLC.ActDestinationIONumber = 0;
            PLC.ActMultiDropChannelNumber = 0;
            PLC.ActThroughNetworkType = 0;
            //PLC.ActDestinationPortNumber = 5002;
            PLC.ActHostAddress = IP_PLC.Text;
            PLC.ActTimeOut = 6000;
        }
       
        public bool PLC_Status()
        {
            return true;
        }

        public void thietlap(string IP_PLC)
        {
            PLC_bond.ActUnitNumber = 26;
            PLC_bond.ActNetworkNumber = 1;
            PLC_bond.ActStationNumber = 5;
            PLC_bond.ActUnitNumber = 0;
            PLC_bond.ActConnectUnitNumber = 0;
            PLC_bond.ActIONumber = 1023;
            PLC_bond.ActCpuType = 35;
            PLC_bond.ActSourceNetworkNumber = 1;
            PLC_bond.ActSourceStationNumber = 2;
            PLC_bond.ActDestinationIONumber = 0;
            PLC_bond.ActMultiDropChannelNumber = 0;
            PLC_bond.ActThroughNetworkType = 0;
            PLC_bond.ActDestinationPortNumber = 5002;
            PLC_bond.ActHostAddress = IP_PLC;
            PLC_bond.ActTimeOut = 60000;
        }
       
    }
}
