using ExcelDna.Integration.CustomUI;
using ExcelDna_MVVM.Utils;
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelDna_MVVM.GUI
{
    public class TaskPaneService
    {
        #region Logger
        private static Logger logger = LogManager.GetCurrentClassLogger();
        private dynamic application;
        private Func<int> getHwnd;
        Func<dynamic> getWindow;
        #endregion

        #region ctor
        public TaskPaneService(dynamic application, Func<int> getHwnd, Func<dynamic> getWindow)
        {
            this.application = application;
            this.getHwnd = getHwnd;
            this.getWindow = getWindow;
        }
        #endregion

        #region Properties & Variables
        private Dictionary<string, CustomTaskPane> panes = new Dictionary<string, CustomTaskPane>();
        #endregion

        #region Public Functions
        public void ShowInTaskPane(string id, string header, bool multiContent, UIElement child, MsoCTPDockPosition dockPostion, MsoCTPDockPositionRestrict positionRestict, double height = -1, double width = -1, Action<bool> VisibleStateChangedAction = null)
        {
            try
            {
                if (child != null && !panes.ContainsKey(id))
                {
                    CustomTaskPane newPane = null;
                    int hwnd = getHwnd();
                    var container = new ucwfWPFContainer();
                    logger.Info($"ShowInTaskPane: id={id}, hwnd={hwnd}, header={header}, height={height}, width={width} " + container.ToString());

                    object parent = getWindow();
                    if (parent != null)
                    {
                        newPane = CustomTaskPaneFactory.CreateCustomTaskPane(container, header, parent);
                    }
                    else
                    {
                        newPane = CustomTaskPaneFactory.CreateCustomTaskPane(container, header);
                    }
                            (newPane.ContentControl as ucwfWPFContainer).Child = child;
                    newPane.DockPosition = dockPostion;

                    if (height != -1)
                        newPane.Height = (int)(height * Win32Helper.GetDpiYScale);

                    if (width != -1)
                        newPane.Width = (int)(width * Win32Helper.GetDpiXScale);

                    newPane.DockPositionRestrict = positionRestict;

                    newPane.VisibleStateChange += (pane) =>
                    {
                        VisibleStateChangedAction?.Invoke(pane.Visible);
                    };

                    newPane.Visible = true;
                    panes.Add(id, newPane);
                }
                if (panes.ContainsKey(id))
                {
                    if (child != null)
                        (panes[id].ContentControl as ucwfWPFContainer).Child = child;
                    panes[id].Visible = child != null;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }
        public void SetTaskPaneVisisbleState(string id, bool state)
        {
            if (panes.ContainsKey(id))
            {
                panes[id].Visible = state;
            }
        }
        #endregion
    }
}
