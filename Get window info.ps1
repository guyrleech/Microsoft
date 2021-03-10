<#
    Get window information - title, window handle and parent's, process and thread ids, position and size

    @guyrleech 2021
#>

[CmdletBinding()]

Param
(
)

## modified from https://stackoverflow.com/questions/25771386/get-all-windows-of-a-process-in-powershell http://reinventingthewheel.azurewebsites.net/MainWindowHandleIsALie.aspx

$TypeDef = @'

using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Api
{
 public class WinStruct
 {
   public string Title {get; set; }
   public IntPtr Hwnd { get; set; }
   public IntPtr ParentHwnd { get; set; }
   public int Left { get; set; }
   public int Top { get; set; }
   public int Width { get; set; }
   public int Height { get; set; }
   public uint PID { get; set; }
   public uint TID { get; set; }
 }

 public class ApiDef
 {
   private delegate bool CallBackPtr(int hwnd, int lParam);
   private static CallBackPtr callBackPtr = Callback;
   private static List<WinStruct> _WinStructList = new List<WinStruct>();
   
   [StructLayout(LayoutKind.Sequential)]

   public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

   [DllImport("user32.dll", SetLastError = true)]
   private static extern bool EnumWindows(CallBackPtr lpEnumFunc, IntPtr lParam);

   [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
   static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

   [DllImport("user32.dll", SetLastError = true)]
   static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
   
   [DllImport("user32.dll", SetLastError = true)]
   static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

   [DllImport("user32.dll", SetLastError=true)]
   static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); 
   
   [DllImport("user32.dll", SetLastError=true)]
   static extern IntPtr GetParent(IntPtr hWnd); 

   private static bool Callback(int hWnd, int lparam)
   {
       StringBuilder sb = new StringBuilder(256);
       int res = GetWindowText((IntPtr)hWnd, sb, 256);
       RECT clientRect = new RECT();
       GetClientRect( (IntPtr)hWnd , out clientRect);
       RECT windowRect = new RECT();
       GetWindowRect( (IntPtr)hWnd , out windowRect);
       uint pid = 0 , tid = 0 ;
       tid = GetWindowThreadProcessId( (IntPtr)hWnd , out pid );
       IntPtr hParent = GetParent( (IntPtr)hWnd ) ;

       _WinStructList.Add(new WinStruct { Hwnd = (IntPtr)hWnd, Title = sb.ToString(), Left = windowRect.Left , 
            Top = windowRect.Top, Width = clientRect.Right - clientRect.Left , Height = clientRect.Bottom - clientRect.Top , PID = pid , TID = tid , ParentHwnd = hParent });
       return true;
   }   

   public static List<WinStruct> GetWindows()
   {
      _WinStructList = new List<WinStruct>();
      EnumWindows(callBackPtr, IntPtr.Zero);
      return _WinStructList;
   }
 }
}
'@

Add-Type -TypeDefinition $TypeDef -ErrorAction Stop

[Api.Apidef]::GetWindows()
