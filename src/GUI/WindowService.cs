namespace ExcelDna_MVVM.GUI
{
    using ExcelDna_MVVM.Utils;
    #region Usings
    using NLog;
    using System;
    using System.Windows.Controls;
    #endregion

    public class WindowService
    {
        #region LoggerInit
        private static Logger logger = LogManager.GetCurrentClassLogger();
        #endregion

        #region Porperties & Variables
        public int Hwnd { get; internal set; }
        public double RibbonWidth { get; internal set; }
        public double RibbonHeight { get; internal set; }
        #endregion

        public void ShowOverlay()
        {

            var rct = Win32Helper.GetParentWindowSize(this, new IntPtr(Hwnd));
            var tlc = new System.Drawing.Point(0, 0);
            tlc.X = (int)rct.Left + (int)((RibbonWidth * Win32Helper.GetDpiXScale) / 3.0);//(app_scriptedit.CommandBars["Ribbon"].Width
            tlc.Y = (int)(rct.Top + RibbonHeight * Win32Helper.GetDpiYScale);//app_scriptedit.CommandBars["Ribbon"].Height


            var wnd = new wdwpfOverlayDialog()
            {
                TopLeftCorner = new System.Windows.Point(tlc.X, tlc.Y),
                HeightMaximized = true,
                WidthMaximized = false,
                ChildUserControl = new System.Windows.Controls.UserControl() { Content = new TextBlock() { Text = "Hallo Welt" } },
                Width = 500,
                ParentHwnd = new IntPtr(Hwnd),
            };
        }
    }
}
