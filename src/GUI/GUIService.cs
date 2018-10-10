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
        public Dispatcher Dispatcher { get; set; }
        public TaskPaneService TaskPaneService { get; set; }
        #endregion

        #region public Functions
        public void ShowOverlay(UserControl content, double verticalOffset = 0, double horizontalOffset = 0)
        {
            var wnd = GetOverlayWindow(content, verticalOffset, horizontalOffset);
            wnd.ShowDialog();
        }
        public Window GetOverlayWindow(UserControl content, double verticalOffset = 0, double horizontalOffset = 0)
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
        public void ShowWaitingControl(object content, Action<bool> VisibleStateChangedAction = null)
        {
            TaskPaneService.ShowInTaskPane("WaitingPane"
                , (string)(LocalizeDictionary.Instance.GetLocalizedObject("se-xll:SenseExcelRibbon:StatusPaneHeader", null, LocalizeDictionary.Instance.Culture))
                , false
                , content as UIElement
                , ExcelDna.Integration.CustomUI.MsoCTPDockPosition.msoCTPDockPositionBottom
                , ExcelDna.Integration.CustomUI.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange
                , 80
                , VisibleStateChangedAction: VisibleStateChangedAction);
        }
        public void ShowSelectionControl(object content, Action<bool> VisibleStateChangedAction = null)
        {
            TaskPaneService.ShowInTaskPane("SelectionControl"
                , "Sense"
                , false
                , content as UIElement
                , ExcelDna.Integration.CustomUI.MsoCTPDockPosition.msoCTPDockPositionTop
                , ExcelDna.Integration.CustomUI.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange
                , 80
                , VisibleStateChangedAction: VisibleStateChangedAction);
        }
        #endregion

        #region private Functions        
        #endregion
    }
}
