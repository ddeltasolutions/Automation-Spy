using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Interop.UIAutomationClient;

namespace dDeltaSolutions.Spy
{
    internal class Helper
    {
        public static string ControlTypeIdToString(int id)
        {
            if (id == UIA_ControlTypeIds.UIA_ButtonControlTypeId)
            {
                return "UIA_ButtonControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_CalendarControlTypeId)
            {
                return "UIA_CalendarControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_CheckBoxControlTypeId)
            {
                return "UIA_CheckBoxControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_ComboBoxControlTypeId)
            {
                return "UIA_ComboBoxControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_EditControlTypeId)
            {
                return "UIA_EditControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_HyperlinkControlTypeId)
            {
                return "UIA_HyperlinkControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_ImageControlTypeId)
            {
                return "UIA_ImageControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_ListItemControlTypeId)
            {
                return "UIA_ListItemControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_ListControlTypeId)
            {
                return "UIA_ListControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_MenuControlTypeId)
            {
                return "UIA_MenuControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_MenuBarControlTypeId)
            {
                return "UIA_MenuBarControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_MenuItemControlTypeId)
            {
                return "UIA_MenuItemControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_ProgressBarControlTypeId)
            {
                return "UIA_ProgressBarControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_RadioButtonControlTypeId)
            {
                return "UIA_RadioButtonControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_ScrollBarControlTypeId)
            {
                return "UIA_ScrollBarControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_SliderControlTypeId)
            {
                return "UIA_SliderControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_SpinnerControlTypeId)
            {
                return "UIA_SpinnerControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_StatusBarControlTypeId)
            {
                return "UIA_StatusBarControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_TabControlTypeId)
            {
                return "UIA_TabControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_TabItemControlTypeId)
            {
                return "UIA_TabItemControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_TextControlTypeId)
            {
                return "UIA_TextControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_ToolBarControlTypeId)
            {
                return "UIA_ToolBarControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_ToolTipControlTypeId)
            {
                return "UIA_ToolTipControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_TreeControlTypeId)
            {
                return "UIA_TreeControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_TreeItemControlTypeId)
            {
                return "UIA_TreeItemControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_CustomControlTypeId)
            {
                return "UIA_CustomControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_GroupControlTypeId)
            {
                return "UIA_GroupControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_ThumbControlTypeId)
            {
                return "UIA_ThumbControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_DataGridControlTypeId)
            {
                return "UIA_DataGridControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_DataItemControlTypeId)
            {
                return "UIA_DataItemControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_DocumentControlTypeId)
            {
                return "UIA_DocumentControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_SplitButtonControlTypeId)
            {
                return "UIA_SplitButtonControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_WindowControlTypeId)
            {
                return "UIA_WindowControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_PaneControlTypeId)
            {
                return "UIA_PaneControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_HeaderControlTypeId)
            {
                return "UIA_HeaderControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_HeaderItemControlTypeId)
            {
                return "UIA_HeaderItemControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_TableControlTypeId)
            {
                return "UIA_TableControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_TitleBarControlTypeId)
            {
                return "UIA_TitleBarControlTypeId";
            }
            else if (id == UIA_ControlTypeIds.UIA_SeparatorControlTypeId)
            {
                return "UIA_SeparatorControlTypeId";
            }
            
            return "";
        }
        
        [DllImport("oleacc.dll")]
        public static extern uint GetStateText(uint dwStateBit, [Out] StringBuilder lpszStateBit, uint cchStateBitMax);
        
        public static Dictionary<uint, string> statesDict = null;
        
        public static string GetStatesAsText(uint state)
        {
            if (state == 0)
            {
                return "normal";
            }
            
            if (statesDict == null)
            {
                statesDict = new Dictionary<uint, string>();
                InitStates();
            }
            
            string states = "";
            foreach (uint crtState in statesDict.Keys)
            {
                if ((state & crtState) != 0)
                {
                    if (states != "")
                    {
                        states += ",";
                    }
                    states += statesDict[crtState];
                }
            }
            
            return states;
        }
        
        private static void InitStates()
        {
            statesDict.Add(0x0, "normal");
            statesDict.Add(0x40000000, "haspopup");
            statesDict.Add(0x1, "unavailable");
            statesDict.Add(0x2, "selected");
            statesDict.Add(0x4, "focused");
            statesDict.Add(0x8, "pressed");
            statesDict.Add(0x10, "checked");
            statesDict.Add(0x20, "mixed");
            statesDict.Add(0x40, "readonly");
            statesDict.Add(0x80, "hottracked");
            statesDict.Add(0x100, "default");
            statesDict.Add(0x200, "expanded");
            statesDict.Add(0x400, "collapsed");
            statesDict.Add(0x800, "busy");
            statesDict.Add(0x1000, "floating");
            statesDict.Add(0x2000, "marqueed");
            statesDict.Add(0x4000, "animated");
            statesDict.Add(0x8000, "invisible");
            statesDict.Add(0x10000, "offscreen");
            statesDict.Add(0x20000, "sizeable");
            statesDict.Add(0x40000, "moveable");
            statesDict.Add(0x80000, "selfvoicing");
            statesDict.Add(0x100000, "focusable");
            statesDict.Add(0x200000, "selectable");
            statesDict.Add(0x400000, "linked");
            statesDict.Add(0x800000, "traversed");
            statesDict.Add(0x1000000, "multiselectable");
            statesDict.Add(0x2000000, "extselectable");
            statesDict.Add(0x4000000, "alert_low");
            statesDict.Add(0x8000000, "alert_medium");
            statesDict.Add(0x10000000, "alert_high");
            statesDict.Add(0x20000000, "protected");
            //statesDict.Add(0x3FFFFFFF, "valid");
        }
        
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool ScreenToClient(IntPtr hwnd, ref POINT lpPoint);
        
        [DllImport("user32.dll")]
        public static extern IntPtr GetDC(IntPtr hWnd);
        
        [DllImport("gdi32.dll", EntryPoint = "CreateCompatibleDC", SetLastError=true)]
        public static extern IntPtr CreateCompatibleDC([In] IntPtr hdc);
        
        [DllImport("gdi32.dll", EntryPoint = "CreateCompatibleBitmap")]
        public static extern IntPtr CreateCompatibleBitmap([In] IntPtr hdc, int nWidth, int nHeight);
        
        [DllImport("gdi32.dll", EntryPoint = "SelectObject")]
        public static extern IntPtr SelectObject([In] IntPtr hdc, [In] IntPtr hgdiobj);
        
        [DllImport("gdi32.dll", EntryPoint = "BitBlt", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool BitBlt([In] IntPtr hdc, int nXDest, int nYDest, int nWidth, int nHeight, 
            [In] IntPtr hdcSrc, int nXSrc, int nYSrc, TernaryRasterOperations dwRop);
            
        [DllImport("gdi32.dll", EntryPoint = "DeleteObject")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool DeleteObject([In] IntPtr hObject);
        
        [DllImport("gdi32.dll", EntryPoint = "DeleteDC")]
        public static extern bool DeleteDC([In] IntPtr hdc);
        
        [DllImport("user32.dll")]
        public static extern bool ReleaseDC(IntPtr hWnd, IntPtr hDC);
        
        //[DllImport("user32.dll")]
        //static extern IntPtr GetWindowDC(IntPtr hWnd);
        
        [DllImport("user32.dll", EntryPoint="GetWindowLong")]
        static extern IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex);
        
        //[DllImport("user32.dll", SetLastError = true)]
        //public static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);
        
        public enum TernaryRasterOperations : uint 
        {
            SRCCOPY     = 0x00CC0020,
            SRCPAINT    = 0x00EE0086,
            SRCAND      = 0x008800C6,
            SRCINVERT   = 0x00660046,
            SRCERASE    = 0x00440328,
            NOTSRCCOPY  = 0x00330008,
            NOTSRCERASE = 0x001100A6,
            MERGECOPY   = 0x00C000CA,
            MERGEPAINT  = 0x00BB0226,
            PATCOPY     = 0x00F00021,
            PATPAINT    = 0x00FB0A09,
            PATINVERT   = 0x005A0049,
            DSTINVERT   = 0x00550009,
            BLACKNESS   = 0x00000042,
            WHITENESS   = 0x00FF0062,
            CAPTUREBLT  = 0x40000000 //only if WinVer >= 5.0.0 (see wingdi.h)
        }
        
        /*[StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;        // x position of upper-left corner
            public int Top;         // y position of upper-left corner
            public int Right;       // x position of lower-right corner
            public int Bottom;      // y position of lower-right corner

            public RECT(int left, int top, int right, int bottom)
            {
                this.Left = left;
                this.Top = top;
                this.Right = right;
                this.Bottom = bottom;
            }
        }*/
        
        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;

            public POINT(int x, int y)
            {
                this.X = x;
                this.Y = y;
            }

            public POINT(System.Drawing.Point pt) : this(pt.X, pt.Y) { }

            public static implicit operator System.Drawing.Point(POINT p)
            {
                return new System.Drawing.Point(p.X, p.Y);
            }

            public static implicit operator POINT(System.Drawing.Point p)
            {
                return new POINT(p.X, p.Y);
            }
        }
        
        //static int GWL_STYLE = -16;
        //static long WS_CHILD = 0x40000000;
        
        /*private static bool IsWindows10OrHigher()
		{
			return ((Environment.OSVersion.Version.Major == 6 && Environment.OSVersion.Version.Minor >= 2) ||
                Environment.OSVersion.Version.Major >= 7);
		}*/
        
        /*private static bool IsTopLevelWindow(IntPtr hWnd)
        {
            IntPtr stylePtr = GetWindowLongPtr(hWnd, GWL_STYLE);
            if (stylePtr == IntPtr.Zero)
            {
                return false;
            }
            long style = stylePtr.ToInt32();
            return ((style & WS_CHILD) == 0);
        }*/
        
        /*private static Bitmap CreateWindowSnapshot(IntPtr hWnd, IUIAutomationElement element)
        {
            //System.Windows.MessageBox.Show("CreateWindowSnapshot");
            
            RECT rect;
            if (GetWindowRect(hWnd, out rect) == false)
            {
                return null;
            }
            
            int width = rect.Right - rect.Left;
            int height = rect.Bottom - rect.Top;
            int left = 0;
            int top = 0;
            
            IUIAutomationTreeWalker tw = MainWindow.uiAutomation.ControlViewWalker;
            IUIAutomationElement root = MainWindow.uiAutomation.GetRootElement();
            //MessageBox.Show(Environment.OSVersion.Version.Major.ToString() + ", " + Environment.OSVersion.Version.Minor);
            if (IsWindows8OrHigher() && element.CurrentControlType == UIA_ControlTypeIds.UIA_WindowControlTypeId && IsTopLevelWindow(hWnd))
            {
                width = width - 16;
                height = height - 9;
                left = 8;
                top = 1;
                //System.Windows.MessageBox.Show("Windows 10");
            }

            IntPtr hdc = GetWindowDC(hWnd);
            //IntPtr hdc = UnsafeNativeFunctions.GetDC(hWnd);
            if (hdc == IntPtr.Zero)
            {
                return null;
            }
            IntPtr hdcMem = CreateCompatibleDC(hdc);
            if (hdcMem == IntPtr.Zero)
            {
                return null;
            }
            IntPtr hbmp = CreateCompatibleBitmap(hdc, width, height);
            if (hbmp == IntPtr.Zero)
            {
                return null;
            }
            if (SelectObject(hdcMem, hbmp) == IntPtr.Zero)
            {
                return null;
            }
            if (BitBlt(hdcMem, 0, 0, width, height, hdc, left, top, 
                TernaryRasterOperations.SRCCOPY | TernaryRasterOperations.CAPTUREBLT) == false)
            {
                return null;
            }

            Bitmap bitmap = System.Drawing.Image.FromHbitmap(hbmp);

            DeleteObject(hbmp);
            DeleteDC(hdcMem);
            ReleaseDC(hWnd, hdc);
            return bitmap;
        }*/
        
        private static Bitmap CreateSnapshot(IntPtr hWnd, int left, int top, int right, int bottom)
        {
            POINT pt;
            pt.X = left;
            pt.Y = top;
            if (ScreenToClient(hWnd, ref pt) == false)
            {
                return null;
            }
            if (pt.X < 0 || pt.Y < 0)
            {
                return null;
            }
            int width = right - left;
            int height = bottom - top;

            //HDC hdc = GetWindowDC(hWnd);
            IntPtr hdc = GetDC(hWnd);
            if (hdc == IntPtr.Zero)
            {
                return null;
            }
            IntPtr hdcMem = CreateCompatibleDC(hdc);
            if (hdcMem == IntPtr.Zero)
            {
                return null;
            }
            IntPtr hbmp = CreateCompatibleBitmap(hdc, width, height);
            if (hbmp == IntPtr.Zero)
            {
                return null;
            }
            if (SelectObject(hdcMem, hbmp) == IntPtr.Zero)
            {
                return null;
            }
            if (BitBlt(hdcMem, 0, 0, width, height, hdc, pt.X, pt.Y, 
                TernaryRasterOperations.SRCCOPY) == false)
            {
                return null;
            }

            //::MessageBox(hwndDlg, "WM_PAINT", "", MB_OK);

            //PBITMAPINFO pInfo = CreateBitmapInfoStruct(hWnd, hbmp);
            //CreateBMPFile(hWnd, (TCHAR*)_T("capture.bmp"), pInfo, hbmp, hdcMem);
            Bitmap bitmap = System.Drawing.Image.FromHbitmap(hbmp);

            DeleteObject(hbmp);
            DeleteDC(hdcMem);
            ReleaseDC(hWnd, hdc);
            return bitmap;
        }
        
        /*private static IUIAutomationElement GetParentWindow(IUIAutomationElement element)
        {
            //IntPtr hWnd = element.CurrentNativeWindowHandle;
            IUIAutomationTreeWalker tw = MainWindow.uiAutomation.ControlViewWalker;
            IUIAutomationElement root = MainWindow.uiAutomation.GetRootElement();
            IUIAutomationElement crtElement = element;
            
            if (crtElement == null || MainWindow.CompareElements(crtElement, root) == true)
            {
                return null;
            }
            
            //while (hWnd == IntPtr.Zero)
            do
            {
                crtElement = tw.GetParentElement(crtElement);
                if (crtElement == null || MainWindow.CompareElements(crtElement, root) == true)
                {
                    return null;
                }
                
                IntPtr hWnd = crtElement.CurrentNativeWindowHandle;
                if (hWnd != IntPtr.Zero)
                {
                    return crtElement;
                }
            }
            while (true);
            
            return null;
        }*/
        
        internal static Bitmap CaptureElementToBitmap(IUIAutomationElement element)
        {
            try
            {
                tagRECT rect = element.CurrentBoundingRectangle;
                
                // remove windows 10 extra space for top level windows ********
                //bool isWindows10OrHigher = (Environment.OSVersion.Version.Major == 6 && Environment.OSVersion.Version.Minor >= 2) ||
                //    Environment.OSVersion.Version.Major >= 7;
                
                //IntPtr hWnd = element.CurrentNativeWindowHandle;
                //if (hWnd != IntPtr.Zero && element.CurrentControlType == UIA_ControlTypeIds.UIA_WindowControlTypeId && 
                //    isWindows10OrHigher && IsTopLevelWindow(hWnd) /*IsWindows10OrHigher()*/ )
                /*{
                    rect.left += 9;
                    rect.top += 1;
                    rect.right -= 9;
                    rect.bottom -= 9;
                }*/
                // ************************************************************
                
                IntPtr hWndDesktop = MainWindow.uiAutomation.GetRootElement().CurrentNativeWindowHandle;
                return CreateSnapshot(hWndDesktop, rect.left, rect.top, rect.right, rect.bottom);
            }
            catch 
            {
                return null;
            }
            
            /*if (element.CurrentControlType == 50009 || 
                element.CurrentControlType == 50011)
            {
                return null;
            }
            
            if (element.CurrentControlType == 50016 && element.CurrentFrameworkId == "WinForm")
            {
                element = MainWindow.uiAutomation.ControlViewWalker.GetParentElement(element);
            }
            
            try
            {
                Bitmap bitmap = null;
                IntPtr hWnd = element.CurrentNativeWindowHandle;
                if (hWnd != IntPtr.Zero)
                {
                    bitmap = CreateWindowSnapshot(hWnd, element);
                    if (bitmap != null)
                    {
                        return bitmap;
                    }
                }
                IUIAutomationElement crtParent = element;
                if (hWnd == IntPtr.Zero)
                {
                    crtParent = GetParentWindow(element);
                    if (crtParent == null)
                    {
                        return null;
                    }
                    hWnd = crtParent.CurrentNativeWindowHandle;
                }
                
                tagRECT rect = element.CurrentBoundingRectangle;
                
                while (bitmap == null)
                {
                    bitmap = CreateSnapshot(hWnd, rect.left, rect.top, rect.right, rect.bottom);
                    if (bitmap == null)
                    {
                        crtParent = GetParentWindow(crtParent);
                        if (crtParent == null)
                        {
                            //System.Windows.MessageBox.Show("maximize");
                            //return null;
                            hWnd = MainWindow.uiAutomation.GetRootElement().CurrentNativeWindowHandle;
                            return CreateSnapshot(hWnd, rect.left, rect.top, rect.right, rect.bottom);
                        }
                        hWnd = crtParent.CurrentNativeWindowHandle;
                    }
                }
                
                return bitmap;
            }
            catch
            {
                return null;
            }*/
        }
        
        public static void CaptureElementToFile(/*IUIAutomationElement element*/
            Bitmap bitmap, string fileName)
        {
            /*if (element.CurrentControlType == 50009 || 
                element.CurrentControlType == 50011)
            {
                return;
            }
            
            //element.SetFocus();
            
            if (element.CurrentControlType == 50016 && element.CurrentFrameworkId == "WinForm")
            {
                element = MainWindow.uiAutomation.ControlViewWalker.GetParentElement(element);
            }*/
            
            /*Bitmap bitmap = null;
            IntPtr hWnd = element.CurrentNativeWindowHandle;
            if (hWnd != IntPtr.Zero)
            {
                bitmap = CreateWindowSnapshot(hWnd, element);
                if (bitmap == null)
                {
                    IUIAutomationElement crtParent = element;
                    if (hWnd == IntPtr.Zero)
                    {
                        crtParent = GetParentWindow(element);
                        if (crtParent == null)
                        {
                            return;
                        }
                        hWnd = crtParent.CurrentNativeWindowHandle;
                    }
                    
                    tagRECT rect = element.CurrentBoundingRectangle;
                    
                    while (bitmap == null)
                    {
                        bitmap = CreateSnapshot(hWnd, rect.left, rect.top, rect.right, rect.bottom);
                        if (bitmap == null)
                        {
                            crtParent = GetParentWindow(crtParent);
                            if (crtParent == null)
                            {
                                return;
                            }
                            hWnd = crtParent.CurrentNativeWindowHandle;
                        }
                    }
                }
            }*/
            
            /*Bitmap bitmap = null;
            IntPtr hWnd = element.CurrentNativeWindowHandle;
            if (hWnd != IntPtr.Zero)
            {
                bitmap = CreateWindowSnapshot(hWnd, element);
                if (bitmap != null)
                {
                    return bitmap;
                }
            }
            IUIAutomationElement crtParent = element;
            if (hWnd == IntPtr.Zero)
            {
                crtParent = GetParentWindow(element);
                if (crtParent == null)
                {
                    return null;
                }
                hWnd = crtParent.CurrentNativeWindowHandle;
            }
            
            tagRECT rect = element.CurrentBoundingRectangle;
            
            while (bitmap == null)
            {
                bitmap = CreateSnapshot(hWnd, rect.left, rect.top, rect.right, rect.bottom);
                if (bitmap == null)
                {
                    crtParent = GetParentWindow(crtParent);
                    if (crtParent == null)
                    {
                        return null;
                    }
                    hWnd = crtParent.CurrentNativeWindowHandle;
                }
            }*/
            
            /*Bitmap bitmap = CaptureElementToBitmap(element);
            if (bitmap == null)
            {
                System.Windows.MessageBox.Show("Cannot capture element");
                return;
            }*/
            
            try
            {
                ImageFormat format = ImageFormat.Bmp;
                string extension = System.IO.Path.GetExtension(fileName).ToLower();
                if (extension == "jpg" || extension == "jpeg")
                {
                    format = ImageFormat.Jpeg;
                }
                else if (extension == "png")
                {
                    format = ImageFormat.Png;
                }
                
                bitmap.Save(fileName, format);
            }
            catch 
            {
                System.Windows.MessageBox.Show("Cannot save image");
            }
            bitmap.Dispose();
        }
    }
    
    enum Modes
    {
        Control,
        Raw,
        Content
    }

    internal class Traces
    {
        private static string traceFile = "trace.txt";

        private static bool enableTrace = true;

        public static void Trace(string message)
        {
#if DEBUG
            if (false == Traces.enableTrace)
            {
                return;
            }

            string line = DateTime.Now.ToString(/*"G"*/) + ": " + message + Environment.NewLine;

            try
            {
                File.AppendAllText(Traces.traceFile, line);
            }
            catch { }
#endif
        }

        public static string TraceFile
        {
            get
            {
                return Traces.traceFile;
            }
        }

        public static void ClearTrace()
        {
#if DEBUG
            if (false == Traces.enableTrace)
            {
                return;
            }

            try
            {
                if (File.Exists(Traces.traceFile))
                {
                    File.Delete(Traces.traceFile);
                }
            }
            catch { }
#endif
        }
    }

    public class TreeNode
    {
        private IUIAutomationElement internalElement = null;
        private bool isRootElement = false;

        public TreeNode(IUIAutomationElement el, bool isRoot = false)
        {
            this.internalElement = el;
            this.isRootElement = isRoot;
        }
        
        public bool IsRoot
        {
            get
            {
                return isRootElement;
            }
        }

        public bool IsAlive
        {
            get
            {
                int processId = 0;
                try
                {
                    processId = this.internalElement.CurrentProcessId;
                }
                catch
                {}

                return (processId != 0);
            }
        }

        public override string ToString()
        {
            string returnString = string.Empty;
            
            if (this.internalElement == null)
            {
                return "";
            }

            try
            {
                /*if (this.isRootElement == true)
                {
                    returnString = "Desktop (" + this.internalElement.Current.LocalizedControlType + ")";
                }
                else
                {*/
                    /*if (internalElement == null)
                    {
                        System.Windows.MessageBox.Show("null element");
                    }*/
                    string name = this.internalElement.CurrentName;

                    if (name == null)
                    {
                        name = "";
                    }
                    else if (name.Length > 30)
                    {
                        name = name.Substring(0, 30);
                        name += "...";
                    }
                    
                    //string localizedControlType = this.internalElement.CurrentLocalizedControlType;
					string localizedControlType = null;
					int controlType = this.internalElement.CurrentControlType;
                    //if (string.IsNullOrEmpty(localizedControlType))
                    //{
					if (cacheTypes.ContainsKey(controlType))
					{
						localizedControlType = cacheTypes[controlType];
					}
					else
					{
                        localizedControlType = Helper.ControlTypeIdToString(controlType);
                        //localizedControlType = localizedControlType.Replace("UIA_", "").Replace("ControlTypeId", "").ToLower();
						if (localizedControlType != "")
						{
							//localizedControlType = localizedControlType.Remove(0, 4).Replace("ControlTypeId", "");
							localizedControlType = localizedControlType.Remove(0, 4);
							localizedControlType = localizedControlType.Remove(localizedControlType.Length - 13);
						}
						
						cacheTypes.Add(controlType, localizedControlType);
						
                        /*if (this.internalElement.CurrentControlType != UIA_ControlTypeIds.UIA_CustomControlTypeId)
                        {
                            System.Windows.MessageBox.Show("null or empty");
                        }*/
                    }

                    name = "\"" + name + "\"";
                    returnString = name + " (" + localizedControlType + ")";
                //}
            }
            catch (System.Exception ex)
            { 
                //System.Windows.MessageBox.Show(ex.Message);
            }

            return returnString;
        }
		
		private static Dictionary<int, string> cacheTypes = new Dictionary<int, string>();

        public IUIAutomationElement Element
        {
            get
            {
                return this.internalElement;
            }
        }

        public List<TreeNode> Children
        {
            get
            {
                List<TreeNode> childrenList = new List<TreeNode>();
                List<IUIAutomationElement> children = MainWindow.FindChildren(this.internalElement);

                foreach (IUIAutomationElement element in children)
                {
                    TreeNode node = new TreeNode(element);
                    childrenList.Add(node);
                }

                return childrenList;
            }
        }
    }
    
    public enum AccRoles
    {
        ROLE_SYSTEM_ALERT = 8,
        ROLE_SYSTEM_ANIMATION = 54,
        ROLE_SYSTEM_APPLICATION = 14,
        ROLE_SYSTEM_BORDER = 19,
        ROLE_SYSTEM_BUTTONDROPDOWN = 56,
        ROLE_SYSTEM_BUTTONDROPDOWNGRID = 58,
        ROLE_SYSTEM_BUTTONMENU = 57,
        ROLE_SYSTEM_CARET = 7,
        ROLE_SYSTEM_CELL = 29,
        ROLE_SYSTEM_CHARACTER = 32,
        ROLE_SYSTEM_CHART = 17,
        ROLE_SYSTEM_CHECKBUTTON = 44,
        ROLE_SYSTEM_CLIENT = 10,
        ROLE_SYSTEM_CLOCK = 61,
        ROLE_SYSTEM_COLUMN = 27,
        ROLE_SYSTEM_COLUMNHEADER = 25,
        ROLE_SYSTEM_COMBOBOX = 46,
        ROLE_SYSTEM_CURSOR = 6,
        ROLE_SYSTEM_DIAGRAM = 53,
        ROLE_SYSTEM_DIAL = 49,
        ROLE_SYSTEM_DIALOG = 18,
        ROLE_SYSTEM_DOCUMENT = 15,
        ROLE_SYSTEM_DROPLIST = 47,
        ROLE_SYSTEM_EQUATION = 55,
        ROLE_SYSTEM_GRAPHIC = 40,
        ROLE_SYSTEM_GRIP = 4,
        ROLE_SYSTEM_GROUPING = 20,
        ROLE_SYSTEM_HELPBALLOON = 31,
        ROLE_SYSTEM_HOTKEYFIELD = 50,
        ROLE_SYSTEM_INDICATOR = 39,
        ROLE_SYSTEM_IPADDRESS = 63,
        ROLE_SYSTEM_LINK = 30,
        ROLE_SYSTEM_LIST = 33,
        ROLE_SYSTEM_LISTITEM = 34,
        ROLE_SYSTEM_MENUBAR = 2,
        ROLE_SYSTEM_MENUITEM = 12,
        ROLE_SYSTEM_MENUPOPUP = 11,
        ROLE_SYSTEM_OUTLINE = 35,
        ROLE_SYSTEM_OUTLINEBUTTON = 64,
        ROLE_SYSTEM_OUTLINEITEM = 36,
        ROLE_SYSTEM_PAGETAB = 37,
        ROLE_SYSTEM_PAGETABLIST = 60,
        ROLE_SYSTEM_PANE = 16,
        ROLE_SYSTEM_PROGRESSBAR = 48,
        ROLE_SYSTEM_PROPERTYPAGE = 38,
        ROLE_SYSTEM_PUSHBUTTON = 43,
        ROLE_SYSTEM_RADIOBUTTON = 45,
        ROLE_SYSTEM_ROW = 28,
        ROLE_SYSTEM_ROWHEADER = 26,
        ROLE_SYSTEM_SCROLLBAR = 3,
        ROLE_SYSTEM_SEPARATOR = 21,
        ROLE_SYSTEM_SLIDER = 51,
        ROLE_SYSTEM_SOUND = 5,
        ROLE_SYSTEM_SPINBUTTON = 52,
        ROLE_SYSTEM_SPLITBUTTON = 62,
        ROLE_SYSTEM_STATICTEXT = 41,
        ROLE_SYSTEM_STATUSBAR = 23,
        ROLE_SYSTEM_TABLE = 24,
        ROLE_SYSTEM_TEXT = 42,
        ROLE_SYSTEM_TITLEBAR = 1,
        ROLE_SYSTEM_TOOLBAR = 22,
        ROLE_SYSTEM_TOOLTIP = 13,
        ROLE_SYSTEM_WHITESPACE = 59,
        ROLE_SYSTEM_WINDOW = 9
    }
}