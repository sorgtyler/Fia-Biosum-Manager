using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
    public partial class uc_filesize_monitor : UserControl
    {
        private System.Timers.Timer timerCheckForFileSize;
        private string _strPath = "";
        private string _strFullPath = "";
        private int _intMaxSize;
        private Color _oProgressBarColor;
        private string _strInfo = "";
        
        public uc_filesize_monitor()
        {
            InitializeComponent();
            _oProgressBarColor = progressBarBasic1.ForeColor;
            timerCheckForFileSize = new System.Timers.Timer();
            timerCheckForFileSize.Interval = 5000; //every 5 seconds
            timerCheckForFileSize.Elapsed += timerCheckForFileSize_Tick;
            timerCheckForFileSize.Stop();
        }

        void timerCheckForFileSize_Tick(Object source, System.Timers.ElapsedEventArgs e)
        {
            if (System.IO.File.Exists(_strFullPath))
            {
                System.IO.FileInfo fi = new System.IO.FileInfo(this._strFullPath);
                long size = (long)fi.Length;
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Label)lblSize, "Text", GetSizeReadable(size));
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Label)lblSize, "Left", (int)(this.ClientSize.Width * .5) - (int)(this.lblSize.Width * .5));
                UpdateThermPercent(1, _intMaxSize, (int)size);
            }
            



        }
        public void BeginMonitoringFile(
            string p_strFile,
            int p_intMaxSize,
            string p_strMaxSizeLabel)
        {
            frmMain.g_oDelegate.SetControlPropertyValue(progressBarBasic1, "Value", 0);
            _strFullPath = p_strFile;
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Label)lblFileName, "Text", frmMain.g_oUtils.getFileName(p_strFile));
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Label)lblFileName, "Left", (int)(this.ClientSize.Width * .5) - (int)(this.lblFileName.Width * .5));
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Label)lblMaxSize, "Text", p_strMaxSizeLabel);
            progressBarBasic1.ForeColor = _oProgressBarColor;
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Label)lblSize, "Text", "");
            _intMaxSize = p_intMaxSize;
            timerCheckForFileSize.Enabled = true;
            timerCheckForFileSize.Start();
            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.UserControl)this, "Show");
            

           
            
            
            
        }
        public void EndMonitoringFile()
        {
            timerCheckForFileSize.Stop();
            this._strFullPath = "";
            this._strInfo = "";
            this._strPath = "";
            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.UserControl)this, "Hide");
        }
        private void UpdateThermPercent(int p_intMin, int p_intMax, int p_intValue)
        {
            int intPercent = (int)(((double)(p_intValue - p_intMin) /
                (double)(p_intMax - p_intMin)) * 100);

            if (intPercent >= 90)
                frmMain.g_oDelegate.SetControlPropertyValue(progressBarBasic1, "ForeColor", Color.Red);
                
            frmMain.g_oDelegate.SetControlPropertyValue(progressBarBasic1, "Value", intPercent);


        }
        public int CurrentPercent(string p_strFullPath, int p_intMax)
        {
            int intPercent=0;
            if (System.IO.File.Exists(_strFullPath))
            {
                System.IO.FileInfo fi = new System.IO.FileInfo(p_strFullPath);
                long currentsize = (long)fi.Length;
                intPercent = (int)(((double)(currentsize) /
                  (double)(p_intMax)) * 100);
            }

            return intPercent;

        }
        // Returns the human-readable file size for an arbitrary, 64-bit file size
        //  The default format is "0.### XB", e.g. "4.2 KB" or "1.434 GB"
        private string GetSizeReadable(long i)
        {
            string sign = (i < 0 ? "-" : "");
            double readable = (i < 0 ? -i : i);
            string suffix;
            if (i >= 0x1000000000000000) // Exabyte
            {
                suffix = "EB";
                readable = (double)(i >> 50);
            }
            else if (i >= 0x4000000000000) // Petabyte
            {
                suffix = "PB";
                readable = (double)(i >> 40);
            }
            else if (i >= 0x10000000000) // Terabyte
            {
                suffix = "TB";
                readable = (double)(i >> 30);
            }
            else if (i >= 0x40000000) // Gigabyte
            {
                suffix = "GB";
                readable = (double)(i >> 20);
            }
            else if (i >= 0x100000) // Megabyte
            {
                suffix = "MB";
                readable = (double)(i >> 10);
            }
            else if (i >= 0x400) // Kilobyte
            {
                suffix = "KB";
                readable = (double)i;
            }
            else
            {
                return i.ToString(sign + "0 B"); // Byte
            }
            readable = readable / 1024;

            return sign + readable.ToString("0.### ") + suffix;
        }
       
        public string  Information
        {
            get { return _strInfo; }
            set { _strInfo = value; }
        }
        public string File
        {
            get { return _strFullPath; }
        }
        private void btnInfo_Click(object sender, EventArgs e)
        {
            string str = "";
            str = "File Name\r\n";
            str = str + "----------\r\n";
            str = str + frmMain.g_oUtils.getFileName(_strFullPath) + "\r\n\r\n";
            str = str + "Location\r\n";
            str = str + "----------\r\n";
            str = str + frmMain.g_oUtils.getDirectory(_strFullPath) + "\r\n\r\n";
            if (_strInfo.Trim().Length > 0)
            {
                str = str + "Other Information\r\n";
                str = str + "--------------------\r\n";
                str = str + _strInfo;
            }
            MessageBox.Show(str,"FIA Biosum");

        }

    }
}
