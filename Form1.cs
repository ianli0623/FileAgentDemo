using log4net;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace FileAgentDemo
{
    public partial class Form1 : Form
    {
        private ILog g_logger = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        //全局變量
        //SharpClipboard clipboard = new SharpClipboard();

        [DllImport("User32.dll")]
        public extern static bool AddClipboardFormatListener(IntPtr hwnd);
        // See http://msdn.microsoft.com/en-us/library/ms649021%28v=vs.85%29.aspx
        public const int WM_CLIPBOARDUPDATE = 0x031D;
        public const int WM_PASTE = 0x302;
        public static IntPtr HWND_MESSAGE = new IntPtr(-3);

        public delegate void MyInvoke(string str1);

        public Form1()
        {
            InitializeComponent();
            AddClipboardFormatListener(this.Handle);
            ProcessCount("FileAgent");
        }

        private void Start_Click(object sender, EventArgs e)
        {
            if (Start.Text == "取消")
            {
                FolderPath.Enabled = true;
                Start.Text = "鎖定";
            }
            else
            {
                if (!string.IsNullOrEmpty(FolderPath.Text))
                {
                    FolderPath.Enabled = false;
                    Start.Text = "取消";
                    MyFileSystemWatcher(FolderPath.Text);
                }
                else
                {
                    MessageBox.Show("請輸入資料夾路徑");
                }
                if (!System.IO.Directory.Exists(FolderPath.Text))
                {
                    MessageBox.Show("請輸入已存在資料夾");
                }
            }
        }
        private void MyFileSystemWatcher(string m_Path)
        {
            try
            {
                FileSystemWatcher _watch = new FileSystemWatcher();

                _watch.Path = m_Path;

                g_logger.InfoFormat("Monitor Path: {0}", _watch.Path);

                //設定所要監控的變更類型
                //_watch.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.Size;
                _watch.NotifyFilter = NotifyFilters.DirectoryName | NotifyFilters.FileName | NotifyFilters.Size;

                //設定所要監控的檔案
                _watch.Filter = "*.*";

                //設定是否監控子資料夾
                _watch.IncludeSubdirectories = true;

                //設定是否啟動元件，此部分必須要設定為 true，不然事件是不會被觸發的
                _watch.EnableRaisingEvents = true;

                //設定觸發事件
                g_logger.Info("File Monitor Start...");
                _watch.Created += new FileSystemEventHandler(_watch_Created);
                _watch.Changed += new FileSystemEventHandler(_watch_Changed);
                ////伺服器版
                _watch.Renamed += new RenamedEventHandler(_watch_Renamed);
                _watch.Deleted += new FileSystemEventHandler(_watch_Deleted);
            }
            catch (Exception ex)
            {
                g_logger.Error(ex.Message, ex);
            }
        }

        #region  監聽事件
        /// <summary>
        /// 當所監控的資料夾有建立文字檔時觸發
        /// </summary>
        private void _watch_Created(object sender, FileSystemEventArgs e)
        {
            try
            {
                DirectoryInfo dirInfo = new DirectoryInfo(e.FullPath.ToString());
                FileInfo newFile = new FileInfo(dirInfo.ToString());
                string Content_text = "行為:新增" + Environment.NewLine
                    + "電腦名稱:" + GetServerName() + Environment.NewLine
                    + "登入帳號:" + GetLoginAccount() + Environment.NewLine
                    + "登入IP:" + GetIpAddresses().TrimEnd('、') + Environment.NewLine
                    + "事件資料夾:" + dirInfo.FullName.Replace(dirInfo.Name, "") + Environment.NewLine
                    + "新增檔名:" + dirInfo.Name + Environment.NewLine
                    + "建立時間:" + dirInfo.CreationTime.ToString("yyyy/MM/dd HH:mm:ss") + Environment.NewLine;
                MyInvoke mi = new MyInvoke(UpdateForm);
                this.BeginInvoke(mi, new Object[] { Content_text });

            }
            catch (UnauthorizedAccessException ex)
            {
                //g_logger.Info("[File Unauthorized!] " + ex.Message, ex);
            }
        }

        /// <summary>
        /// 當所監控的資料夾有文字檔檔案內容有異動時觸發
        /// </summary>
        private void _watch_Changed(object sender, FileSystemEventArgs e)
        {
            try
            {
                DirectoryInfo dirInfo = new DirectoryInfo(e.FullPath.ToString());
                FileInfo newFile = new FileInfo(dirInfo.ToString());
                MyInvoke mi = new MyInvoke(UpdateForm);
                string Content_text = "行為:修改" + Environment.NewLine
                    + "電腦名稱:" + GetServerName() + Environment.NewLine
                    + "登入帳號:" + GetLoginAccount() + Environment.NewLine
                    + "登入IP:" + GetIpAddresses().TrimEnd('、') + Environment.NewLine
                    + "事件資料夾:" + dirInfo.FullName.Replace(dirInfo.Name, "") + Environment.NewLine
                    + "異動檔名:" + dirInfo.Name + Environment.NewLine
                    + "異動內容時間:" + dirInfo.LastWriteTime.ToString("yyyy/MM/dd HH:mm:ss") + Environment.NewLine;
                this.BeginInvoke(mi, new Object[] { Content_text });
            }
            catch (UnauthorizedAccessException ex)
            {
                //g_logger.Info("[File Unauthorized!] " + ex.Message, ex);
            }
        }

        /// <summary>
        /// 當所監控的資料夾有文字檔檔案重新命名時觸發
        /// </summary>
        private void _watch_Renamed(object sender, RenamedEventArgs e)
        {
            try
            {
                DirectoryInfo dirInfo = new DirectoryInfo(e.FullPath.ToString());
                DirectoryInfo dirInfo_old = new DirectoryInfo(e.OldFullPath.ToString());
                MyInvoke mi = new MyInvoke(UpdateForm);
                string Content_text = "行為:重新命名" + Environment.NewLine
                    + "電腦名稱:" + GetServerName() + Environment.NewLine
                    + "登入帳號:" + GetLoginAccount() + Environment.NewLine
                    + "登入IP:" + GetIpAddresses().TrimEnd('、') + Environment.NewLine
                    + "事件資料夾:" + dirInfo.FullName.Replace(dirInfo.Name, "") + Environment.NewLine
                    + "更名前:" + dirInfo_old.Name.ToString() + Environment.NewLine
                    + "更名後:" + dirInfo.Name.ToString() + Environment.NewLine
                    + "更名前路徑:" + e.OldFullPath.ToString() + Environment.NewLine
                    + "更名後路徑:" + e.FullPath.ToString() + Environment.NewLine;
                this.BeginInvoke(mi, new Object[] { Content_text });
            }
            catch (UnauthorizedAccessException)
            {
            }
        }

        /// <summary>
        /// 當所監控的資料夾有文字檔檔案有被刪除時觸發
        /// </summary>
        private void _watch_Deleted(object sender, FileSystemEventArgs e)
        {
            try
            {
                DirectoryInfo dirInfo = new DirectoryInfo(e.FullPath.ToString());
                MyInvoke mi = new MyInvoke(UpdateForm);
                string Content_text = "行為:刪除" + Environment.NewLine
                    + "電腦名稱:" + GetServerName() + Environment.NewLine
                    + "登入帳號:" + GetLoginAccount() + Environment.NewLine
                    + "登入IP:" + GetIpAddresses().TrimEnd('、') + Environment.NewLine
                    + "事件資料夾:" + dirInfo.FullName.Replace(dirInfo.Name, "") + Environment.NewLine
                    + "被刪除檔名:" + dirInfo.Name.ToString() + Environment.NewLine
                    + "被刪除檔名路徑:" + e.FullPath.ToString() + Environment.NewLine;
                this.BeginInvoke(mi, new Object[] { Content_text });
            }
            catch (UnauthorizedAccessException)
            {
            }
        }
        #endregion

        #region  環境參數
        /// <summary>
        /// Gets the name of the server.
        /// </summary>
        /// <returns>Server name</returns>
        public static string GetServerName()
        {
            return Environment.MachineName;
        }

        /// <summary>
        /// Gets the login account.
        /// </summary>
        /// <returns>Login account</returns>
        public static string GetLoginAccount()
        {
            return System.Security.Principal.WindowsIdentity.GetCurrent().Name;
        }

        /// <summary>
        /// Gets the ip addresses.
        /// </summary>
        /// <returns>ip addresses</returns>
        public static string GetIpAddresses()
        {
            //取IP另一種方法
            string IP4Address = "";
            List<IPAddress> m_IPs = Dns.GetHostAddresses(Dns.GetHostName()).Where(a => a.AddressFamily.ToString() == "InterNetwork").ToList();
            foreach (IPAddress IPA in m_IPs)
            {
                if (IPA.ToString().Contains("10."))
                {
                    IP4Address = IPA.ToString();
                    break;
                }
            }
            if (string.IsNullOrEmpty(IP4Address))
            {
                int m_count = m_IPs.Count();
                IP4Address = m_IPs[m_count - 1].ToString();
            }
            return IP4Address;
        }

        #endregion
        /// <summary>
        ///  檢查處理程序筆數
        /// </summary>
        /// <param name="ProcessName"></param>
        /// <returns></returns>
        private int ProcessCount(string ProcessName)
        {
            int m_result = 0;
            try
            {
                foreach (var process in Process.GetProcessesByName(ProcessName))
                {
                    m_result++;
                }

                if (m_result > 1)
                {
                    g_logger.Error(m_result);
                    foreach (var process in Process.GetProcessesByName(ProcessName))
                    {
                        process.Kill();
                        m_result--;
                        if (m_result == 1)
                        {
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                g_logger.Error(ex);
            }
            return m_result;
        }

        public void UpdateForm(string param1)
        {
            this.Content.Text = param1;
        }
    }
}
