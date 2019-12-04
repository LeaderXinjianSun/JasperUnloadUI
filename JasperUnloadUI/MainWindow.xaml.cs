using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using JasperUnloadUI.Model;
using BingLibrary.hjb.file;
using System.Diagnostics;
using OfficeOpenXml;
using System.IO;
using System.Data;

namespace JasperUnloadUI
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        #region 变量
        Fx5u Fx5u;
        string iniParameterPath = System.Environment.CurrentDirectory + "\\Parameter.ini";
        long SWms = 0;
        Scan Scan5;
        int[] BordIndex = new int[96];
        #endregion
        public MainWindow()
        {
            InitializeComponent();
            System.Diagnostics.Process[] myProcesses = System.Diagnostics.Process.GetProcessesByName("JasperUnloadUI");//获取指定的进程名   
            if (myProcesses.Length > 1) //如果可以获取到知道的进程名则说明已经启动
            {
                System.Windows.MessageBox.Show("不允许重复打开软件");
                System.Windows.Application.Current.Shutdown();
            }
        }
        #region 功能函数
        void AddMessage(string str)
        {
            this.Dispatcher.Invoke(new Action(() => {

                string[] s = MsgTextBox.Text.Split('\n');
                if (s.Length > 1000)
                {
                    MsgTextBox.Text = "";
                }
                if (MsgTextBox.Text != "")
                {
                    MsgTextBox.Text += "\r\n";
                }
                MsgTextBox.Text += DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + " " + str;

            }));
        }
        void Init()
        {
            try
            {
                string ExIoExcelPath = System.Environment.CurrentDirectory + "\\排版.xlsx";

                if (File.Exists(ExIoExcelPath))
                {
                    FileInfo existingFile = new FileInfo(ExIoExcelPath);
                    using (ExcelPackage package = new ExcelPackage(existingFile))
                    {
                        // get the first worksheet in the workbook
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[2];
                        for (int i = 0; i < 96; i++)
                        {
                            BordIndex[i] = int.Parse(worksheet.Cells[i / 12 + 1, i % 12 + 1].Value.ToString());
                        }
                    }
                    string ip = Inifile.INIGetStringValue(iniParameterPath, "FX5U", "Ip", "192.168.0.20");
                    int port = int.Parse(Inifile.INIGetStringValue(iniParameterPath, "FX5U", "Port", "504"));
                    Fx5u = new Fx5u(ip, port);
                    Scan5 = new Scan();
                    string COM = Inifile.INIGetStringValue(iniParameterPath, "Scan", "Scan5", "COM3");
                    Scan5.ini(COM);
                    Run();
                }
                else
                {
                    throw new Exception("排版文件不存在");
                }
            }
            catch (Exception ex)
            {
                AddMessage(ex.Message);
            }
        }
        void CheckBarcode(string barcode,int index)
        {
            if (barcode != "Error")
            {
                Mysql mysql = new Mysql();
                if (mysql.Connect())
                {

                    string stm = "SELECT * FROM BODMSG WHERE SCBODBAR = '" + barcode + "' ORDER BY SIDATE DESC LIMIT 0,5";
                    DataSet ds = mysql.Select(stm);
                    DataTable dt = ds.Tables["table0"];
                    if (dt.Rows.Count > 0)
                    {
                        if (dt.Rows[0]["STATUS"] == DBNull.Value)
                        {
                            AddMessage("板 " + barcode + " 状态栏位为空");
                            switch (index)
                            {
                                case 0:
                                    Fx5u.SetM("M2505", true);
                                    break;
                                case 1:
                                    Fx5u.SetM("M2507", true);
                                    break;
                                default:
                                    break;
                            }                            
                        }
                        else
                        {
                            if ((string)dt.Rows[0]["STATUS"] == "OFF")
                            {
                                AddMessage("板 " + barcode + " 是未测板");
                                switch (index)
                                {
                                    case 0:
                                        Fx5u.SetM("M2505", true);
                                        break;
                                    case 1:
                                        Fx5u.SetM("M2507", true);
                                        break;
                                    default:
                                        break;
                                }
                            }
                            else
                            {
                                stm = "INSERT INTO BODMSG (SCBODBAR, STATUS) VALUES('" + barcode + "','OFF')";
                                AddMessage("板 " + barcode + " 解绑");
                                mysql.executeQuery(stm);

                                stm = "SELECT * FROM barbind WHERE SCBODBAR = '" + barcode + "' ORDER BY SIDATE DESC LIMIT 0,96";
                                ds = mysql.Select(stm);
                                dt = ds.Tables["table0"];

                                if (dt.Rows.Count == 96)
                                {
                                    //string datetimestr = (string)dt.Rows[0]["SIDATE"];
                                    short[] result = new short[96];
                                    bool checkrst = true;
                                    for (int i = 0; i < 96; i++)
                                    {
                                        DataRow[] drs = dt.Select(string.Format("PCSSER = '{0}'", (BordIndex[i]).ToString()));
                                        if (drs.Length == 1)
                                        {
                                            try
                                            {
                                                result[i] = short.Parse((string)drs[0]["RESULT"]);
                                            }
                                            catch (Exception ex)
                                            {
                                                AddMessage(ex.Message);
                                                checkrst = false;
                                                break;
                                            }
                                        }
                                        else
                                        {
                                            AddMessage("板 " + barcode + " 序号 " + (BordIndex[i]).ToString() + "索引数 " + drs.Length.ToString());
                                            checkrst = false;
                                            break;
                                        }
                                    }
                                    if (checkrst)
                                    {
                                        string str;
                                        switch (index)
                                        {
                                            case 0:
                                                Fx5u.WriteMultD("D1000", result);
                                                str = "A_BordInfo;";
                                                for (int i = 0; i < 96; i++)
                                                {
                                                    str += result[i].ToString() + ";";
                                                }
                                                str = str.Substring(0, str.Length - 1);
                                                AddMessage(str);
                                                Fx5u.SetM("M2504", true);
                                                break;
                                            case 1:
                                                Fx5u.WriteMultD("D1100", result);
                                                str = "B_BordInfo;";
                                                for (int i = 0; i < 96; i++)
                                                {
                                                    str += result[i].ToString() + ";";
                                                }
                                                str = str.Substring(0, str.Length - 1);
                                                AddMessage(str);
                                                Fx5u.SetM("M2506", true);
                                                break;
                                            default:
                                                break;
                                        }
                                    }
                                    else
                                    {
                                        switch (index)
                                        {
                                            case 0:
                                                Fx5u.SetM("M2505", true);
                                                break;
                                            case 1:
                                                Fx5u.SetM("M2507", true);
                                                break;
                                            default:
                                                break;
                                        }
                                    }
                                }
                                else
                                {
                                    AddMessage("板 " + barcode + " 产品信息条目 " + dt.Rows.Count.ToString() + " < 96");
                                    switch (index)
                                    {
                                        case 0:
                                            Fx5u.SetM("M2505", true);
                                            break;
                                        case 1:
                                            Fx5u.SetM("M2507", true);
                                            break;
                                        default:
                                            break;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        AddMessage("板 " + barcode + " 信息未录入");
                        switch (index)
                        {
                            case 0:
                                Fx5u.SetM("M2505", true);
                                break;
                            case 1:
                                Fx5u.SetM("M2507", true);
                                break;
                            default:
                                break;
                        }
                    }
                }
                else
                {
                    AddMessage("Mysql数据库查询失败");
                    switch (index)
                    {
                        case 0:
                            Fx5u.SetM("M2505", true);
                            break;
                        case 1:
                            Fx5u.SetM("M2507", true);
                            break;
                        default:
                            break;
                    }
                }
                mysql.DisConnect();
            }
            else
            {
                switch (index)
                {
                    case 0:
                        Fx5u.SetM("M2505", true);
                        break;
                    case 1:
                        Fx5u.SetM("M2507", true);
                        break;
                    default:
                        break;
                }
            }
        }
        async void Run()
        {
            bool first = false;
            bool[] M2000;
            bool m2002 = false, m2003 = false;
            Stopwatch sw = new Stopwatch();
            while (true)
            {
                sw.Restart();
                await Task.Delay(100);
                #region UpdateUI
                if (Fx5u.Connect)
                {
                    EllipsePLCState.Fill = Brushes.Green;
                }
                else
                {
                    EllipsePLCState.Fill = Brushes.Red;
                }
                CycleText.Text = SWms.ToString() + " ms";
                #endregion
                #region Work
                M2000 = await Task.Run<bool[]>(()=> {
                    return Fx5u.ReadMultiM("M2000", 16);
                });
                if (M2000 != null)
                {
                    if (first)
                    {
                        first = false;
                        m2002 = M2000[2];
                        m2003 = M2000[3];
                    }
                    if (m2002 != M2000[2])
                    {
                        m2002 = M2000[2];
                        if (m2002)
                        {
                            Fx5u.SetM("M2002", false);
                            Fx5u.SetM("M2004", false);
                            Fx5u.SetM("M2005", false);

                            Scan5.GetBarCode((string barcode) =>
                            {
                                AddMessage("下料A扫码:" + barcode);
                                CheckBarcode(barcode, 0);
                            });

                        }
                    }
                    if (m2003 != M2000[3])
                    {
                        m2003 = M2000[3];
                        if (m2003)
                        {
                            Fx5u.SetM("M2003", false);
                            Fx5u.SetM("M2006", false);
                            Fx5u.SetM("M2007", false);

                            Scan5.GetBarCode((string barcode) =>
                            {
                                AddMessage("下料B扫码:" + barcode);
                                CheckBarcode(barcode, 1);
                            });

                        }
                    }
                }
                #endregion
                SWms = sw.ElapsedMilliseconds;
            }
        }
        #endregion
        private void MsgTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            MsgTextBox.ScrollToEnd();
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Init();
            AddMessage("软件加载完成");
        }
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
        }

        private void FreeBordBarcodeButtonClick(object sender, RoutedEventArgs e)
        {
            if (BordBarcode.Text != "")
            {
                try
                {
                    Mysql mysql = new Mysql();
                    if (mysql.Connect())
                    {
                        string stm = "SELECT * FROM BODMSG WHERE SCBODBAR = '" + BordBarcode.Text + "' ORDER BY SIDATE DESC LIMIT 0,5";
                        DataSet ds = mysql.Select(stm);
                        DataTable dt = ds.Tables["table0"];
                        if (dt.Rows.Count > 0)
                        {
                            if ((string)dt.Rows[0]["STATUS"] == "OFF")
                            {
                                AddMessage("板 " + BordBarcode.Text + " 已处于解绑状态");
                            }
                            else
                            {
                                stm = "INSERT INTO BODMSG (SCBODBAR, STATUS) VALUES('" + BordBarcode.Text + "','OFF')";
                                mysql.executeQuery(stm);
                                AddMessage("板 " + BordBarcode.Text + " 解绑");
                                BordBarcode.Text = "";
                            }
                        }
                        else
                        {
                            AddMessage("板 " + BordBarcode.Text + " 信息无记录");
                        }
                    }
                    mysql.DisConnect();
                }
                catch (Exception ex)
                {
                    AddMessage(ex.Message);
                }
                
            }    
        }

        private void 扫码AClick(object sender, RoutedEventArgs e)
        {
            try
            {
                Scan5.GetBarCode(AddMessage);
            }
            catch
            {

                
            }
        }
    }
}
