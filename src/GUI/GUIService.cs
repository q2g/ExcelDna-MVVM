namespace ExcelDna_MVVM.GUI
{
    #region Usings
    using ExcelDna_MVVM.Utils;
    using NLog;
    using System;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Interop;
    using System.Windows.Threading;
    using WPFLocalizeExtension.Engine;
    #endregion

    public class GUIService
    {
        #region LoggerInit
        private static Logger logger = LogManager.GetCurrentClassLogger();
        #endregion

        #region Porperties & Variables        
        public virtual double RibbonWidth { get; set; }
        public virtual double RibbonHeight { get; set; }
        public virtual Func<int> GetHwnd { get; set; }
        public virtual Func<dynamic> GetCurrentWindow { get; set; }
        public Dispatcher Dispatcher { get; set; }
        public TaskPaneService TaskPaneService { get; set; }
        #endregion

        #region public Functions
        public void ShowOverlay(UserControl content, double verticalOffset = 0, double horizontalOffset = 0)
        {
            var wnd = GetOverlayWindow(content, verticalOffset, horizontalOffset);
            wnd.ShowDialog();
        }
        public Window GetOverlayWindow(UserControl content, double verticalOffset = 0, double horizontalOffset = 0, bool heightMaximized = true, bool widthmaximized = true)
        {

            dynamic currwnd = GetCurrentWindow();
            int offset = (currwnd?.WindowState ?? 0) == -4137 ? 8 : 0;
            var rct = Win32Helper.GetParentWindowSize(this, new IntPtr(GetHwnd()));
            var tlc = new System.Drawing.Point(0, 0)
            {
                X = (int)(rct.Left + horizontalOffset),
                Y = (int)(rct.Top + (offset + RibbonHeight + verticalOffset) * Win32Helper.GetDpiYScale)
            };


            var wnd = new wdwpfOverlayDialog()
            {
                TopLeftCorner = new Point(tlc.X, tlc.Y),
                ChildUserControl = content,
                HeightMaximized = heightMaximized,
                WidthMaximized = widthmaximized,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ParentHwnd = new IntPtr(GetHwnd())
            };
            return wnd;
        }
        public Window GetWindow()
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

        #region private Functions        
        #endregion
    }
}
