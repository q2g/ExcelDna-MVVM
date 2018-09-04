namespace ExcelDna_MVVM.GUI
{
    using ExcelDna_MVVM.Utils;
    #region Usings
    using NLog;
    using System;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Interop;
    #endregion

    public class WindowService
    {
        #region LoggerInit
        private static Logger logger = LogManager.GetCurrentClassLogger();
        #endregion

        #region Porperties & Variables        
        public virtual double RibbonWidth { get; set; }
        public virtual double RibbonHeight { get; set; }
        public virtual Func<int> GetHwnd { get; set; }
        #endregion

        #region public Functions
        public virtual void ShowOverlay(UserControl content, double verticalOffset = 0, double horizontalOffset = 0)
        {

            var rct = Win32Helper.GetParentWindowSize(this, new IntPtr(GetHwnd()));
            var tlc = new System.Drawing.Point(0, 0)
            {
                X = (int)(rct.Left + horizontalOffset),
                Y = (int)(rct.Top + RibbonHeight * Win32Helper.GetDpiYScale + verticalOffset)
            };


            var wnd = new wdwpfOverlayDialog()
            {
                TopLeftCorner = new System.Windows.Point(tlc.X, tlc.Y),
                ChildUserControl = content,
                HeightMaximized = true,
                WidthMaximized = true,
                Width = 500,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ParentHwnd = new IntPtr(GetHwnd()),
            };
            wnd.ShowDialog();
        }

        public virtual Window GetWindow()
        {
            try
            {
                var wnd = new Window()
                {
                };
                var hwnd = -1;
                if (GetHwnd != null)
                    hwnd = GetHwnd();
                if (hwnd != -1)
                    new WindowInteropHelper(wnd).Owner = new IntPtr(hwnd);
                return wnd;
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return null;
        }
        #endregion
    }
}
