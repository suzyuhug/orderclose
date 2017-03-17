using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.IO;
using System.Drawing.Printing;

namespace 入库工具
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
     
        private void button1_Click(object sender, EventArgs e)
        {
            errorlog("准备打开FlexFlow主程序", 0);
            string FFpath = $"{Application.StartupPath.ToString()}\\flexflow\\WOSTG\\Flexflow_Client.exe";
            OpenFF(FFpath,0);
                     
          
        }
        private void OpenFF(string FFpath,int id)//启动FlexFlow程序
        {
            if (File.Exists(FFpath))
            {
                System.Diagnostics.Process.Start(FFpath);
                errorlog("打开flexflow程序", 0);
                bool bl = false;
                int hwnd = 0;
                for (int i = 0; i < 20; i++)
                {
                    Thread.Sleep(1000);
                    hwnd = ApiHelper.FindWindowHwnd("WindowsForms10.Window.8.app.0.11c7a8c", "User Login");
                    if (hwnd != 0)
                    {
                        errorlog("获取到登录页面的句柄", 0);
                        bl = true;
                        break;
                    }
                    else
                    {
                        errorlog("获取登录页面句柄失败", 1);

                    }
                }
                if (bl)
                {
                    loginfrom(hwnd,id);

                }
                else
                {
                    Close();
                }
            }
            else
            {
                MessageBox.Show("FlexFlow程序不存在，请联系我！", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            }
        }


        private void loginfrom(int hwnd, int id)//flexflow登录页面
        {
            Thread.Sleep(1000);
            List<int> findhwndlist = ApiHelper.FindControl((IntPtr)hwnd, "WindowsForms10.EDIT.app.0.11c7a8c", null);
            if (findhwndlist.Count > 0)
            {
                ApiHelper.SendMessage(findhwndlist[0], ApiHelper.WM_SETTEXT, 0, Properties.Settings.Default.username);//密码
                errorlog("输入用户名", 0);
                ApiHelper.SendMessage(findhwndlist[1], ApiHelper.WM_SETTEXT, 0, Properties.Settings.Default.password);//用户名
                errorlog("输入密码", 0);
                int bi = ApiHelper.FindWindowEx((IntPtr)hwnd, "WindowsForms10.BUTTON.app.0.11c7a8c", "Login");
                errorlog("获取登录按钮的句柄", 0);
                if (bi != 0)
                {
                    ApiHelper.SendMessage(bi, ApiHelper.WM_CLICK, 0, "");
                    errorlog("点击登录按钮", 0);

                    bool bl = false;
                    int hwnd1 = 0;
                    for (int i = 0; i < 20; i++)
                    {
                        Thread.Sleep(1000);
                        hwnd1 = ApiHelper.FindWindowHwnd("WindowsForms10.Window.8.app.0.11c7a8c", "FF Client 2.8.24");
                        errorlog("等待主窗体的打开", 0);
                        if (hwnd1 != 0)
                        {
                            errorlog("主窗体打开，获取到句柄", 0);
                            bl = true;
                            break;
                        }
                    }

                    if (bl)
                    {
                        if (id == 0)
                        {
                            listorder(hwnd1);
                        }
                        else if (id == 1)
                        {
                            orderlistclose(hwnd1);
                        }
                    }
                    else
                    {
                        errorlog("主窗体打开失败", 1);
                        Close();
                    }
                }
                else
                {
                    errorlog("登录按钮句柄获取失败", 1);
                    Close();
                }

            }
            else
            {
                errorlog("登录窗体句柄获取失败", 1);
                Close();
            }
        }
        private void orderopen(int hwnd, string order, int ordernum)//工单号输入
        {
            Thread.Sleep(1000);
            errorlog("获取到窗体句柄"+hwnd.ToString(), 0);
            List<int> findhwndlist0 = ApiHelper.FindControl((IntPtr)hwnd, "WindowsForms10.Window.8.app.0.11c7a8c", null);
            if (findhwndlist0[2] != 0)
            {
                errorlog("获取到窗体句柄" + findhwndlist0[2].ToString(), 0);
                List<int> findhwndlist1 = ApiHelper.FindControl((IntPtr)findhwndlist0[2], "WindowsForms10.Window.8.app.0.11c7a8c", null);
                if (findhwndlist1[0] != 0)
                {
                    errorlog("获取到窗体句柄" + findhwndlist1[0].ToString(), 0);
                    List<int> findhwndlist2 = ApiHelper.FindControl((IntPtr)findhwndlist1[0], "WindowsForms10.Window.8.app.0.11c7a8c", null);
                    if (findhwndlist2[0] != 0)
                    {
                        errorlog("获取到窗体句柄" + findhwndlist2[0].ToString(), 0);
                        int formhwnd = ApiHelper.FindWindowEx((IntPtr)findhwndlist2[0], "WindowsForms10.Window.8.app.0.11c7a8c", "Form1");
                        if (formhwnd != 0)
                        {
                            errorlog("获取到窗体句柄" + formhwnd.ToString(), 0);
                            int combohwnd = ApiHelper.FindWindowEx((IntPtr)formhwnd, "WindowsForms10.COMBOBOX.app.0.11c7a8c", null);
                            errorlog("获取到输入空的句柄", 0);
                            if (combohwnd != 0)
                            {
                                errorlog("获取到下拉框句柄" + combohwnd.ToString(), 0);
                                int edithwnd = ApiHelper.FindWindowEx((IntPtr)combohwnd, "Edit", null);
                                if (edithwnd != 0)
                                {
                                    errorlog("获取到输入框的句柄" + edithwnd.ToString(), 0);
                                    Thread.Sleep(500);
                                    errorlog("等待500毫秒", 0);
                                    ApiHelper.SendMessage(edithwnd, ApiHelper.WM_SETTEXT, 0, order);
                                    errorlog("输入工单号：" + order, 0);
                                    ApiHelper.SendMessage(edithwnd, ApiHelper.WM_KEYDOWN, 27, null);
                                    ApiHelper.SendMessage(edithwnd, ApiHelper.WM_CHAR, 27, null);
                                    ApiHelper.SendMessage(edithwnd, ApiHelper.WM_KEYUP, 27, null);
                                    errorlog("等待500毫秒", 0);
                                    Thread.Sleep(500);
                                    int but = ApiHelper.FindWindowEx((IntPtr)formhwnd, "WindowsForms10.BUTTON.app.0.11c7a8c", "OK");
                                    if (but != 0)
                                    {
                                        errorlog("获取到OK键的句柄" + but.ToString(), 0);
                                        ApiHelper.SendMessage(but, ApiHelper.WM_CLICK, 0, "");
                                        errorlog("点击OK键", 0);
                                        if (messageview())
                                        {
                                            for (int i = 0; i < ordernum; i++)
                                            {
                                                errorlog("准备输入工单的数量", 0);
                                                num(findhwndlist2[0], order);
                                            }
                                        }
                                        else
                                        {
                                            errorlog(order+"入库失败", 1);
                                        }                                       
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

      



        private void listorder(int hwnd)//加载工单号
        {
            errorlog("正在查询要入库的工单", 0);
            string sqlstr = "SELECT OrderID,QTY FROM Order_Close WHERE Status='Open'";
            DataSet ds = new DataSet();
            ds = SqlHelper.ExcuteDataSet(sqlstr);
            if (ds.Tables[0].Rows.Count >0)
            {
                string order;int qtynum;
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    errorlog("准备入库第"+ i + "个工单", 0);
                    order = ds.Tables[0].Rows[i]["orderID"].ToString();
                    qtynum =int.Parse( ds.Tables[0].Rows[i]["QTY"].ToString());
                    orderopen(hwnd, order,qtynum);

                }
            }
            closefrom(hwnd);
            opencloseorder();
        }

        private void opencloseorder()//打开关闭工单Flexflow程序
        {
            errorlog("准备打开completedPO FlexFlow程序", 0);
            Thread.Sleep(2000); 
            string FFpath = $"{Application.StartupPath.ToString()}\\flexflow\\CompletePO\\Flexflow_Client.exe";
            OpenFF(FFpath,1);

        }


        private void closefrom(int hwnd)//关闭flexflow主窗体
        {
            errorlog("关闭FlexFlow程序", 0);
            ApiHelper.SendMessage(hwnd, ApiHelper.WM_DESTROY, 0, null);          
        }

        private void num(int hwnd,string order)//获取工单数量
        {
            errorlog("获取到窗体句柄" + hwnd.ToString(), 0);
            Thread.Sleep(1000);
            if (hwnd !=0)
            {
                List<int> findhwndlist0 = ApiHelper.FindControl((IntPtr)hwnd, "WindowsForms10.Window.8.app.0.11c7a8c", null);
                if (findhwndlist0[1]!=0)
                {
                    errorlog("获取到窗体句柄" + findhwndlist0[1].ToString(), 0);
                    List<int> findhwndlist = ApiHelper.FindControl((IntPtr)findhwndlist0[1], "WindowsForms10.Window.8.app.0.11c7a8c", null);
                    if (findhwndlist[0]!=0)
                    {
                        errorlog("获取到窗体句柄" + findhwndlist[0].ToString(), 0);
                        int texthwnd = ApiHelper.FindWindowEx((IntPtr)findhwndlist[0], "WindowsForms10.EDIT.app.0.11c7a8c", null);

                        if (texthwnd!=0)
                        {
                            errorlog("获取到编辑框句柄" + texthwnd.ToString(), 0);
                            ApiHelper.PostMessage(texthwnd, ApiHelper.WM_CHAR, 13, null);
                            errorlog("发送回车键命令", 0);
                            sessmessage(order);                                                        
                        }                        
                    }                   
                }              
            }            
        }
        private void sessmessage(string order)//收工单成功消息框
        {
            Thread.Sleep(1000);
            int Messagehwnd = ApiHelper.FindWindowHwnd("WindowsForms10.Window.8.app.0.11c7a8c", "Flexflow Message : frmPO::ShowSubForm");
            if (Messagehwnd != 0)
            {
                errorlog("获取到正确消息窗体的句柄" + Messagehwnd.ToString(), 0);
                int but = ApiHelper.FindWindowEx((IntPtr)Messagehwnd, "WindowsForms10.BUTTON.app.0.11c7a8c", "OK");
                if (but != 0)
                {
                    errorlog("点击OK按钮", 0);
                    ApiHelper.SendMessage(but, ApiHelper.WM_CLICK, 0, "");
                    upstatus(order, "Active");
                   
                }
            }
        
        }


        private void orderclose(int hwnd,string STG,string order)//工单关闭
        {
            Thread.Sleep(1000);
            if (hwnd!=0)
            {
                errorlog("获取到窗体句柄" + hwnd.ToString(), 0);
                List<int> findhwndlist0 = ApiHelper.FindControl((IntPtr)hwnd, "WindowsForms10.Window.8.app.0.11c7a8c", null);
                if (findhwndlist0[2]!=0)
                {
                    errorlog("获取到窗体句柄" + findhwndlist0[2].ToString(), 0);
                    List<int> findhwndlist1 = ApiHelper.FindControl((IntPtr)findhwndlist0[2], "WindowsForms10.Window.8.app.0.11c7a8c", null);
                    if (findhwndlist1[0]!=0)
                    {
                        errorlog("获取到窗体句柄" + findhwndlist1[0].ToString(), 0);
                        List<int> findhwndlist2 = ApiHelper.FindControl((IntPtr)findhwndlist1[0], "WindowsForms10.Window.8.app.0.11c7a8c", null);
                        if (findhwndlist2[0]!=0)
                        {
                            errorlog("获取到窗体句柄" + findhwndlist2[0].ToString(), 0);
                            List<int> findhwndlist3 = ApiHelper.FindControl((IntPtr)findhwndlist2[0], "WindowsForms10.Window.8.app.0.11c7a8c", null);
                            if (findhwndlist3[1]!=0)
                            {
                                errorlog("获取到窗体句柄" + findhwndlist3[1].ToString(), 0);
                                int texthwnd = ApiHelper.FindWindowEx((IntPtr)findhwndlist3[1], "WindowsForms10.EDIT.app.0.11c7a8c", null);
                                if (texthwnd!=0)
                                {
                                    errorlog("获取到输入框的句柄" + texthwnd.ToString(), 0);
                                    ApiHelper.SendMessage(texthwnd, ApiHelper.WM_SETTEXT, 0, STG);
                                    errorlog("输入号码：" + STG, 0);
                                    Thread.Sleep(500);
                                    errorlog("发送回车键命令", 0);
                                    ApiHelper.PostMessage(texthwnd, ApiHelper.WM_CHAR, 13, null);
                                   
                                    if (messageview())
                                    {
                                        upstatus(order, "Completed");
                                    }
                                }                               
                            }                           
                        }
                    }                   
                }               
            }          
        }


        private void orderlistclose(int hwnd)//加载TERWO###########
        {
            errorlog("正在加载WO号码", 0);
            Thread.Sleep(500);
            string sqlstr = "sp_orderclose";
            DataSet ds = new DataSet();
            ds = SqlHelper.ExcuteDataSet(sqlstr);
            if (ds.Tables[0].Rows.Count > 0)
            {
                string order;string STG;
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    errorlog("正在处理第" + i + "条数据", 0);
                    order = ds.Tables[0].Rows[i]["OrderID"].ToString();
                    STG= ds.Tables[0].Rows[i]["Value"].ToString();
                    orderclose(hwnd, STG,order);

                }
            }

            closefrom(hwnd);
            printorder();

        }


        private static bool messageview()//错误消息框的提示
        {
            
            Thread.Sleep(1000);
            int Messagehwnd = ApiHelper.FindWindowHwnd("WindowsForms10.Window.8.app.0.11c7a8c", "FlexFlow Message");
            if (Messagehwnd != 0)
            {
                int but = ApiHelper.FindWindowEx((IntPtr)Messagehwnd, "WindowsForms10.BUTTON.app.0.11c7a8c", "OK");
                if (but!=0)
                {
                    ApiHelper.SendMessage(but, ApiHelper.WM_CLICK, 0, "");
                    return false;
                }               
            }
            return true;
        }
        private void errorlog(string str,int id)//日志记录
        {
            if (id==0)
            {
                listBox1.Items.Add("正确：" +str);
            }
            else if (id==1)
            {
                listBox1.Items.Add("错误："+str);
            }
            label1.Text = str;
           
            Application.DoEvents();



        }
        private void upstatus(string orderid,string status)//更新状态
        {
            errorlog("更新后台数据为" + status, 0);
            string strsql = $"UPDATE Order_Close SET Status='{status}' WHERE OrderID='{orderid}'";
            if (SqlHelper.ExecuteNonQuery(strsql)) { }          
        }


        private void printorder()//打印入库祥单
        {
            errorlog("准备打印入库明细", 0);
            Thread.Sleep(1000);
            this.printDocument1.DefaultPageSettings.PaperSize = new PaperSize("A4", 2100, 2970);
            this.printDocument1.PrintPage += new PrintPageEventHandler(this.MyPrintDocument_PrintPage);
            PrintController printController = new StandardPrintController();
            printDocument1.PrintController = printController;
            this.printDocument1.Print();
            this.printDocument1.PrintPage -= new PrintPageEventHandler(this.MyPrintDocument_PrintPage);


          


        }
        private void MyPrintDocument_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)//打印的调用        {
            e.Graphics.TranslateTransform(30, 100);
            //e.Graphics.RotateTransform(90.0F);
            string strsql = "exec sp_ordercheck";
            DataSet ds = new DataSet();
            ds = SqlHelper.ExcuteDataSet(strsql);
            if (ds.Tables[0].Rows.Count > 0)
            {
                string str;
                string str1 = "-------------------------------------------------------------------------------------------------------------------------------------------------------------------";
                for (int i = 0; i < ds.Tables[0].Rows.Count ; i++)
                {
                    str = $"     {ds.Tables[0].Rows[i]["编号"].ToString()}   {ds.Tables[0].Rows[i]["料号"].ToString()}   {ds.Tables[0].Rows[i]["ProductionOrderNumber"].ToString()}    {ds.Tables[0].Rows[i]["Value"].ToString()}     {ds.Tables[0].Rows[i]["Description"].ToString()}";
                    errorlog("生成数据：" + str, 0);
                    e.Graphics.DrawString(str, new Font(new FontFamily("Arial"), 10), System.Drawing.Brushes.Black, 0, 25*i);
                    e.Graphics.DrawString(str1, new Font(new FontFamily("Arial"), 10), System.Drawing.Brushes.Black, 0, 25 * i+6);
                }
            }                     
        }

      

        private void Form1_Load(object sender, EventArgs e)
        {
            int ScreenWidth = Screen.PrimaryScreen.WorkingArea.Width;
            int ScreenHeight = Screen.PrimaryScreen.WorkingArea.Height;
            int x = ScreenWidth - this.Width - 5;
            int y = ScreenHeight - this.Height - 5;
            this.Location = new Point(x, y);
        }
    }
}
