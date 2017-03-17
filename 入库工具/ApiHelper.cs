using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace 入库工具
{
    class ApiHelper
    {

        [DllImport("user32", EntryPoint = "GetForegroundWindow")]
        private static extern IntPtr GetForegroundwindow();
        [DllImport("user32.dll", EntryPoint = "FindWindow")]

        

        private static extern IntPtr FindWindow(string IpClassName, string IpWindowName);
        public delegate bool EnumDesktopWindowsDelegate(IntPtr hWnd, uint lParam);
        [DllImport("user32.dll", EntryPoint = "EnumDesktopWindows", ExactSpelling = false, CharSet = CharSet.Auto, SetLastError = true)]

        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumDesktopWindows(IntPtr hDesktop, EnumDesktopWindowsDelegate lpEnumCallbackFunction, IntPtr lParam);

        [DllImport("user32.Dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumChildWindows(IntPtr parentHandle, EnumChildWindowsDelegate callback, IntPtr lParam);
        public delegate bool EnumChildWindowsDelegate(IntPtr hwnd, IntPtr lParam);

        [DllImport("user32.dll", EntryPoint = "FindWindowEx")]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("User32.dll", EntryPoint = "SendMessage")]
        public  static extern int SendMessage(
        int hWnd,             //   handle   to   destination   window
        int Msg,               //   message
        int wParam,       //   first   message   parameter
        string lParam         //   second   message   parameter
        );


        [DllImport("User32.dll", EntryPoint = "PostMessage")]
        public static extern int PostMessage(
       int hWnd,             //   handle   to   destination   window
       int Msg,               //   message
       int wParam,       //   first   message   parameter
       string lParam         //   second   message   parameter
       );
        public const int WM_GETTEXT = 0x000D;
        public const int WM_SETTEXT = 0x000C;
        public const int WM_CLICK = 0x00F5;
        public const int WM_KEYDOWN = 0x0100;
        public const int WM_KEYUP = 0x0101;
        public const int WM_CHAR = 0x0102;
        public const int WM_DRAWITEM = 0x2B;
        public const int WM_CLOSE = 0x0010;
        public const int WM_DESTROY = 0x0002;
        const int WM_SHOWDROPDOWN = 0x014D;//在窗体中声明消息常量
        const int WM_SETCURSOR = 0x14F;
        const int CB_FINDSTRING = 0x14C;



        public static  int FindWindowHwnd(string IpClassName, string IpTitleName)//获取窗体句柄
        {
            if (IpTitleName == "" && IpClassName != "")
            {
                return (int)FindWindow(IpClassName, null);
            }
            else if (IpClassName == "" && IpTitleName != "")
            {
                return (int)FindWindow(null, IpTitleName);
            }
            else if (IpClassName != "" && IpTitleName != "")
            {
                return (int)FindWindow(IpClassName, IpTitleName);
            }
            return 0;
        }

        public  static List<int> FindControl(IntPtr hwnd, string className, string title = null)//获取控件句柄
        {
            List<int> controls = new List<int>();
            IntPtr handle = IntPtr.Zero;
            while (true)
            {
                IntPtr tmp = handle;
                handle = FindWindowEx(hwnd, tmp, className, title);
                if (handle != IntPtr.Zero)
                {
                    controls.Add((int)handle);
                }
                else
                    break;
            }
            return controls;
        }
        public static int FindWindowEx(IntPtr hwnd, string className, string title = null)
        {
            IntPtr handle = IntPtr.Zero;
            IntPtr tmp = handle;
            handle = FindWindowEx(hwnd, tmp , className, title);
             return (int)handle;
           

        }

        public static  int GetForeGroundWindow()//获得顶层窗体的句柄
        {
            return (int)GetForegroundwindow();
        }


    }
}
