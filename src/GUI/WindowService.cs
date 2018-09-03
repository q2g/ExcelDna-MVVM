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
        public virtual int Hwnd { get; set; }
        public virtual double RibbonWidth { get; set; }
        public virtual double RibbonHeight { get; set; }
        public virtual Func<int> GetHwnd { get; set; }
        #endregion

        #region public Functions
        public virtual void ShowOverlay(UserControl content)
        {

            var rct = Win32Helper.GetParentWindowSize(this, new IntPtr(Hwnd));
            var tlc = new System.Drawing.Point(0, 0)
            {
                X = (int)rct.Left + (int)((RibbonWidth * Win32Helper.GetDpiXScale) / 3.0),
                Y = (int)(rct.Top + RibbonHeight * Win32Helper.GetDpiYScale)
            };


            var wnd = new wdwpfOverlayDialog()
            {
                TopLeftCorner = new System.Windows.Point(tlc.X, tlc.Y),
                ChildUserControl = content,
                HeightMaximized = true,
                WidthMaximized = false,
                Width = 500,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ParentHwnd = new IntPtr(Hwnd),
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
