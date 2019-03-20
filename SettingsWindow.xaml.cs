using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace ExcelApplication1
{
    /// <summary>
    /// Settings.xaml 的交互逻辑
    /// </summary>
    public partial class SettingsWindow : Window
    {
        [DllImport("USER32.DLL", CharSet = CharSet.Unicode)]
        private static extern IntPtr GetSystemMenu(IntPtr hWnd, UInt32 bRevert);

        [DllImport("USER32.DLL ", CharSet = CharSet.Unicode)]
        private static extern UInt32 RemoveMenu(IntPtr hMenu, UInt32 nPosition, UInt32 wFlags);

        private const UInt32 SC_CLOSE = 0x0000F060;
        private const UInt32 MF_BYCOMMAND = 0x00000000;

        protected Settings objSettings = new Settings();
        protected bool isCanceled = false;
        public SettingsWindow()
        {
            InitializeComponent();
            
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            this.isCanceled = false;
            this.Close();

        }

        public Settings getSettings()
        {
            return this.objSettings;
        }
        public void setSettings(Settings objSettings)
        {
             this.objSettings =objSettings ;
        }
        private void button2_Click(object sender, RoutedEventArgs e)
        {
            //System.Windows.Forms.FolderBrowserDialog fb = new System.Windows.Forms.FolderBrowserDialog();
            System.Windows.Forms.OpenFileDialog fb = new System.Windows.Forms.OpenFileDialog();
            if (fb.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //选择的文件夹路径
                //txtTargetPath.Text = fb.SelectedPath + "/";
                FileInfo fi = new FileInfo(fb.FileName);
                txtTargetPath.Text = fb.FileName.TrimEnd(fi.Extension.ToCharArray())+".sql";
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (this.isCanceled)
            {
                this.DialogResult = false;
                return;
            }
            this.DialogResult = false;
            e.Cancel = true;
            //if (string.IsNullOrEmpty(this.txtActivityId.Text.Trim()))
            //{
            //    MessageBox.Show("请输入活动ID");
            //    return;
            //}
            if (string.IsNullOrEmpty(this.txtPrizeId.Text.Trim()))
            {
                MessageBox.Show("请输入奖品ID");
                return;
            }
            //if (string.IsNullOrEmpty(this.txtPrjcode.Text.Trim()))
            //{
            //    MessageBox.Show("请输入活动编号");
            //    return;
            //}
            if (string.IsNullOrEmpty(this.txtTargetPath.Text.Trim()))
            {
                MessageBox.Show("请指定保存文件路径");
                return;
            }
            objSettings.ActivityId = this.txtActivityId.Text.Trim();
            objSettings.PrizeId = this.txtPrizeId.Text.Trim();
            objSettings.IsUsed = this.chkIsUsed.IsChecked ?? false;
            objSettings.Prjcode = this.txtPrjcode.Text.Trim();
            objSettings.TargetPath = this.txtTargetPath.Text.Trim();
            objSettings.StartTime = DateTime.Parse(this.txtStartTime.Text).ToString("yyyy-MM-dd HH:mm:ss");
            objSettings.EndTime = DateTime.Parse(this.txtEndTime.Text).ToString("yyyy-MM-dd HH:mm:ss");
            objSettings.Quantity = 0;
            objSettings.CodeLength = 0;
            objSettings.TableIndex = this.cbbTable.SelectedIndex;

            if (objSettings.TableIndex < 0)
            {
                MessageBox.Show("请指定券码表名");
                return;
            }

            e.Cancel = false;
            this.DialogResult = true;
            
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            IntPtr hMenu = GetSystemMenu(hwnd, 0);
            RemoveMenu(hMenu, SC_CLOSE, MF_BYCOMMAND);

            this.isCanceled = false;
            this.txtActivityId.Text = objSettings.ActivityId;
            this.txtPrizeId.Text = objSettings.PrizeId;
            this.chkIsUsed.IsChecked = objSettings.IsUsed ? true : false;
            this.txtPrjcode.Text = objSettings.Prjcode;
            this.txtTargetPath.Text = objSettings.TargetPath;
            var ymd = DateTime.Now.ToString("yyyy-MM-dd");
            if (string.IsNullOrEmpty(objSettings.StartTime)) {
                this.txtStartTime.Text = (ymd + " 00:00:00");
            } else {
                this.txtStartTime.Text = objSettings.StartTime;
            }

            if (string.IsNullOrEmpty(objSettings.EndTime))
            {
                this.txtEndTime.Text = (ymd + " 23:59:59");
            }
            else
            {
                this.txtEndTime.Text = objSettings.EndTime;
            }
            this.cbbTable.SelectedIndex = objSettings.TableIndex;
        }

        private void button3_Click(object sender, RoutedEventArgs e)
        {
            this.isCanceled = true;
            this.Close();
        }
    }
}
