namespace ExcelDna_MVVM.MVVM
{
    #region Usings
    using ExcelDna.Integration;
    using ExcelDna.Integration.CustomUI;
    using ExcelDna_MVVM.Document;
    using ExcelDna_MVVM.Environment;
    using ExcelDna_MVVM.GUI;
    using ExcelDna_MVVM.MVVM.ExcelData;
    using ExcelDna_MVVM.Utils;
    using NetOffice;
    using NetOffice.ExcelApi;
    using NLog;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using System.Threading.Tasks;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Threading;
    using WPFLocalizeExtension.Engine;
    #endregion

    public class MVVMAdapter : IExcelAddIn
    {
        #region Logger
        private static Logger logger = LogManager.GetCurrentClassLogger();
        #endregion

        #region Variables & Properties
        object vmslock = new object();
        private Dictionary<int, List<object>> vms = new Dictionary<int, List<object>>();
        private Dictionary<Type, PropertyInfo> propertyInfos = new Dictionary<Type, PropertyInfo>();
        private Dictionary<Type, List<Type>> vmImplementationTypes = new Dictionary<Type, List<Type>>();
        public Dictionary<int, List<object>> AllVms
        {
            get
            {
                return vms;
            }
        }
        public Dictionary<Type, PropertyInfoCacheItem> servicePropetyInfos = new Dictionary<Type, PropertyInfoCacheItem>();
        object servicePropetyInfosLock = new object();
        //private NetOffice.ExcelApi.Application Application;
        dynamic Application;
        Dispatcher currentDispatcher;
        object sheetID2VMsLock = new object();
        private Dictionary<string, List<object>> sheetID2VMs = new Dictionary<string, List<object>>();
        private CustomTaskPane statusPane;

        #endregion

        #region Events
        public event EventHandler VMCreated;
        public event EventHandler VMDeleted;
        #endregion

        #region IExcelAddIn
        public void AutoClose()
        {
        }

        public void AutoOpen()
        {
            vmImplementationTypes.Add(typeof(IAppVM), TypeUtils.GetTypesImplementingInterface<IAppVM>());
            vmImplementationTypes.Add(typeof(IWorkbookVM), TypeUtils.GetTypesImplementingInterface<IWorkbookVM>());
            vmImplementationTypes.Add(typeof(IWorksheetVM), TypeUtils.GetTypesImplementingInterface<IWorksheetVM>());
            MVVMStatic.Adapter = this;
            currentDispatcher = Dispatcher.CurrentDispatcher;
            //Application = new NetOffice.ExcelApi.Application(null, ExcelDnaUtil.Application);
            Application = ExcelDnaUtil.Application as dynamic;
            var app = new NetOffice.ExcelApi.Application(null, ExcelDnaUtil.Application);
            app.NewWorkbookEvent += Application_NewWorkbookEvent;
            app.WorkbookNewSheetEvent += Application_WorkbookNewSheetEvent;
            app.WorkbookActivateEvent += Application_WorkbookActivateEvent;
            app.SheetActivateEvent += Application_SheetActivateEvent;
            app.WorkbookOpenEvent += App_WorkbookOpenEvent;
            NetOffice.Core.Default.ProxyCountChanged += Default_ProxyCountChanged;
            CreateVMsForApplication(Application);
        }
        #endregion

        #region public Functions       
        #endregion

        #region Eventhandler
        private void Default_ProxyCountChanged(int proxyCount)
        {
            logger.Trace($"ProxyCount Changed value={proxyCount}");
        }
        private void App_WorkbookOpenEvent(Workbook Wb)
        {
            Application_NewWorkbookEvent(Wb);
        }
        private void Application_WorkbookActivateEvent(Workbook wb)
        {
            try
            {
                object objWb = getDynamicWorkbook(wb);
                wb.Dispose();

                dynamic dynWb = objWb as dynamic;
                int aa = dynWb.CustomXMLParts.Count;
                dynamic ii = dynWb.CustomDocumentProperties;

                logger.Trace($"workbook activated {dynWb.Name}");
                RemoveUnusedVms();
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }

        }
        private void Application_SheetActivateEvent(COMObject sheet)
        {

            logger.Trace($"sheet activated {(sheet as Worksheet).Name}");
            RemoveUnusedVms();
            sheet.Dispose();
        }
        private object getDynamicWorkbook(Workbook wb)
        {
            object dynWb = null;
            foreach (var workbook in Application.Workbooks)
            {
                if (workbook.Name == wb.Name)
                {
                    dynWb = workbook;
                    break;
                }
            }
            return dynWb;
        }
        private void Application_NewWorkbookEvent(Workbook wb)
        {
            try
            {
                object dynWb = getDynamicWorkbook(wb);
                wb.Dispose();

                ConvertWorkbookAsync(dynWb).ContinueWith((res) =>
                {
                    if (!res.IsFaulted && res.Result != null)
                    {

                        foreach (var hwnd in res.Result.hwnds)
                        {
                            CreateVMImplementations<IWorkbookVM>(hwnd, res.Result, dynWb);
                            CreateSheetVMsFromWorkbookAsync(res.Result, dynWb);
                        }

                    }
                });

            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }


        }
        private void Application_WorkbookNewSheetEvent(Workbook wb, COMObject sheet)
        {
            try
            {
                object dynWb = getDynamicWorkbook(wb);
                wb.Dispose();
                sheet.Dispose();

                ConvertWorkbookAsync(dynWb).ContinueWith((res) =>
                {
                    if (!res.IsFaulted && res != null)
                        CreateSheetVMsFromWorkbookAsync(res.Result, dynWb);
                });
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }

        }
        #endregion

        #region private Functions
        private Task RemoveUnusedVms()
        {
            Task retval = new Task(() => { });
            try
            {
                List<Task<WorkbookData>> tasks = new List<Task<WorkbookData>>();
                foreach (var wb in Application.Workbooks)
                {
                    tasks.Add(ConvertWorkbookAsync(wb));
                }

                retval = Task.WhenAll(tasks).ContinueWith((res) =>
                {
                    try
                    {
                        List<int> existingHwnds = new List<int>() { -1 };
                        List<string> existingSheetIds = new List<string>();

                        foreach (var task in tasks)
                        {
                            if (!task.IsFaulted)
                            {
                                existingHwnds.AddRange(task.Result.hwnds);
                                var ids = task.Result.sheetIds;
                                if (ids != null)
                                {
                                    existingSheetIds.AddRange(ids);
                                }
                                else
                                {
                                    logger.Warn($"Possible Error: No worksheet for workbook found, aborting {nameof(this.RemoveUnusedVms)}");
                                    return;
                                }
                            }
                            else
                            {
                                logger.Error(task.Exception);
                                return;
                            }
                        }
                        lock (vmslock)
                        {
                            vms.Keys.ToList().Diff(existingHwnds, out var newHwnds, out var removedHwnds);
                            foreach (var hwnd in removedHwnds)
                            {
                                var vmsToRemove = vms[hwnd].ToList();
                                foreach (var vm in vmsToRemove)
                                {
                                    vms[hwnd].Remove(vm);
                                    logger.Trace(() => GetVMsCount());
                                }
                                VMDeleted?.Invoke(this, new VMEventArgs() { VMs = vmsToRemove });
                                foreach (var vm in vmsToRemove)
                                {
                                    if (vm is IDisposable vmToDispose)
                                        vmToDispose.Dispose();
                                    logger.Trace(() => GetVMsCount());
                                }

                                vms.Remove(hwnd);
                            }
                        }

                        lock (sheetID2VMsLock)
                        {
                            sheetID2VMs.Keys.ToList().Diff(existingSheetIds, out var newSheetIds, out var removedSheetIds);
                            var allvms = vms.SelectMany(ele => ele.Value).ToList();
                            foreach (var sheetid in removedSheetIds)
                            {
                                var vmsToRemove = sheetID2VMs[sheetid].ToList();
                                foreach (var vmToRemove in vmsToRemove)
                                {
                                    foreach (var item in vms)
                                    {
                                        if (item.Value.Contains(vmToRemove))
                                        {
                                            item.Value.Remove(vmToRemove);
                                            logger.Trace(() => GetVMsCount());
                                        }
                                    }
                                }
                                VMDeleted?.Invoke(this, new VMEventArgs() { VMs = vmsToRemove });
                                sheetID2VMs.Remove(sheetid);
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex);
                    }
                });
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return retval;
        }
        private void CreateVMsForApplication(dynamic app)
        {
            try
            {
                if (vms == null)
                    vms = new Dictionary<int, List<object>>();

                CreateVMImplementations<IAppVM>(-1, null, null);

                foreach (var workbook in app.Workbooks)
                {
                    Application_NewWorkbookEvent(workbook);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }
        private PropertyInfoCacheItem GetServiceProperetyInfo(Type forType)
        {
            PropertyInfoCacheItem retval = null;
            try
            {
                if (!servicePropetyInfos.ContainsKey(forType))
                {
                    retval = new PropertyInfoCacheItem
                    {
                        PiWindowService = forType.GetProperties().FirstOrDefault(prop => prop.PropertyType.FullName == typeof(WindowService).FullName),
                        PiDocumentPropertyService = forType.GetProperties().FirstOrDefault(prop => prop.PropertyType.FullName == typeof(SeDocument).FullName)
                    };
                }
                else
                {
                    return servicePropetyInfos[forType];
                }
                lock (servicePropetyInfosLock)
                {
                    if (!servicePropetyInfos.ContainsKey(forType) && retval != null)
                    {
                        servicePropetyInfos.Add(forType, retval);
                    }
                }

            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return retval;
        }
        private List<object> CreateVMImplementations<T>(int hwnd, WorkbookData workbookdata, dynamic workbook) where T : IVM
        {
            List<object> createdVms = new List<object>();
            try
            {
                SeDocument documentService = null;
                if (typeof(T) == typeof(IWorkbookVM))
                {
                    documentService = new SeDocument()
                    {
                        Workbook = workbook
                    };
                    documentService.LoadTableJsons();
                }
                var types = vmImplementationTypes[typeof(T)].Where(type => !type.IsAbstract).ToList();
                createdVms = types.Select((type) =>
                {
                    try
                    {
                        logger.Info($"Create VM for Type: {type?.FullName}");
                        var vm = Activator.CreateInstance(type);


                        var servicePropertyInfos = GetServiceProperetyInfo(type);
                        if (servicePropertyInfos != null)
                        {
                            if (servicePropertyInfos.PiWindowService != null)
                            {
                                System.Func<int> GetHwnd = () =>
                                {
                                    var retHwnd = hwnd;
                                    if (retHwnd == -1)
                                        retHwnd = Application?.ActiveWindow?.Hwnd ?? -1;
                                    return retHwnd;
                                };

                                var ws = new WindowService()
                                {
                                    RibbonHeight = Application.CommandBars["Ribbon"].Height,
                                    RibbonWidth = Application.CommandBars["Ribbon"].Width,
                                    Dispatcher = currentDispatcher,
                                    GetHwnd = GetHwnd,
                                    ShowWaitingControl = (control) =>
                                      {
                                          ShowWaitingPane(control as UIElement, GetHwnd());
                                      },
                                    ShowWaitingText = (control) =>
                                      {
                                          currentDispatcher.BeginInvoke((System.Action)(() =>
                                          {
                                              ShowWaitingText(control as string, GetHwnd());
                                          }));

                                      }

                                };
                                servicePropertyInfos.PiWindowService.SetValue(vm, ws);
                            }

                            if (servicePropertyInfos.PiDocumentPropertyService != null && documentService != null)
                            {
                                servicePropertyInfos.PiDocumentPropertyService.SetValue(vm, documentService);
                            }
                        }

                        return vm;
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex);
                    }
                    return null;
                }).ToList();
                Task.Run(() =>
                            {
                                lock (vmslock)
                                {
                                    try
                                    {
                                        foreach (var vm in createdVms)
                                        {
                                            if (vm != null)
                                            {
                                                if (!vms.ContainsKey(hwnd))
                                                    vms.Add(hwnd, new List<object>());
                                                vms[hwnd].Add(vm);
                                                logger.Trace(() => GetVMsCount());
                                            }
                                        }
                                        VMCreated?.Invoke(this, new VMEventArgs() { VMs = createdVms, HWND = hwnd });
                                    }
                                    catch (Exception ex)
                                    {
                                        logger.Error(ex);
                                    }
                                }
                            });
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return createdVms;
        }
        private void ShowWaitingPane(UIElement child, int hwnd)
        {
            if (statusPane == null)
            {
                try
                {
                    var container = new ucwfWPFContainer();
                    logger.Info("Conatiner: " + container.ToString());
                    object parent = null;
                    foreach (var window in Application.Windows)
                    {
                        if (window.Hwnd == hwnd)
                        {
                            parent = window;
                            break;
                        }
                    }
                    if (parent != null)
                    {
                        statusPane = CustomTaskPaneFactory.CreateCustomTaskPane(container, (string)(LocalizeDictionary.Instance.GetLocalizedObject("se-xll:SenseExcelRibbon:StatusPaneHeader", null, LocalizeDictionary.Instance.Culture)), parent);
                    }
                    else
                    {
                        statusPane = CustomTaskPaneFactory.CreateCustomTaskPane(container, (string)(LocalizeDictionary.Instance.GetLocalizedObject("se-xll:SenseExcelRibbon:StatusPaneHeader", null, LocalizeDictionary.Instance.Culture)));
                    }
                    (statusPane.ContentControl as ucwfWPFContainer).Child = child;
                    statusPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionBottom;
                    statusPane.Visible = true;
                    var startHeight = 60.0;
                    if (ExcelDnaUtil.ExcelVersion > 15)
                        startHeight = 80.0;

                    statusPane.Height = (int)(startHeight * Win32Helper.GetDpiYScale);
                    statusPane.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
                }
                catch (Exception ex)
                {
                    logger.Error(ex);
                }
            }


            if (statusPane != null)
            {
                if (child != null)
                {
                    (statusPane.ContentControl as ucwfWPFContainer).Child = child;
                }
                statusPane.Visible = (child != null);
            }
        }
        private void ShowWaitingText(string text, int hwnd)
        {

        }
        private string GetVMsCount()
        {
            var AppVmCount = vms.SelectMany(ele => ele.Value)
                .Select(ele => ele.GetType()).ToList()
                .Count(ele => ele.GetInterfaces().Any(typ => typ.FullName == typeof(IAppVM).FullName));
            var WorkbookVmCount = vms.SelectMany(ele => ele.Value)
                .Select(ele => ele.GetType()).ToList()
                .Count(ele => ele.GetInterfaces().Any(typ => typ.FullName == typeof(IWorkbookVM).FullName));
            var SheetVmCount = vms.SelectMany(ele => ele.Value)
                .Select(ele => ele.GetType()).ToList()
                .Count(ele => ele.GetInterfaces().Any(typ => typ.FullName == typeof(IWorksheetVM).FullName));
            return $"*****************************AppVMs:{AppVmCount}, WorkbookVMs:{WorkbookVmCount}, WorksheetVMs:{SheetVmCount}******************************";
        }
        private Task<WorkbookData> ConvertWorkbookAsync(dynamic wb)
        {
            WorkbookData wbd = new WorkbookData();
            wbd.Name = wb.Name;
            //wbd.hwnds = wb.Windows.Select(win => win.Hwnd).ToList();
            foreach (var window in wb.Windows)
            {
                wbd.hwnds.Add(window.Hwnd);
            }

            return GetSheetIdsFromWorkbookAsync(wbd).ContinueWith((res) =>
            {
                if (!res.IsFaulted && res.Result != null)
                    wbd.sheetIds.AddRange(res.Result);
                return wbd;
            });


        }
        private Task CreateSheetVMsFromWorkbookAsync(WorkbookData wbd, dynamic workbook)
        {
            return Task.Run(() =>
            {
                if (wbd != null)
                {
                    try
                    {
                        var ids = wbd.sheetIds;
                        if (ids != null)
                        {
                            foreach (var hwnd in wbd.hwnds)
                            {
                                foreach (var id in wbd.sheetIds)
                                {
                                    if (!sheetID2VMs.ContainsKey(id))
                                    {//TODO: should there be a worksheet VM for every Window of the WB?
                                     //Is there a case where different Workbookwindows have different sheets?
                                        var impls = CreateVMImplementations<IWorksheetVM>(hwnd, wbd, workbook);
                                        Task.Run(() =>
                                        {
                                            lock (sheetID2VMsLock)
                                            {
                                                if (!sheetID2VMs.ContainsKey(id))
                                                    sheetID2VMs.Add(id, impls);
                                            }
                                        });
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex);
                    }
                }
            });
        }
        private Task<List<string>> GetSheetIdsFromWorkbookAsync(WorkbookData wb)
        {

            return Task.Run(() =>
            {
                var ids = new List<string>();
                try
                {
                    EnvironmentAdapter.QueueAction(() =>
                    {
                        try
                        {
                            var sheetnames = (object[,])XlCall.Excel(XlCall.xlfGetWorkbook, 1, wb.Name);

                            for (int j = 0; j < sheetnames.GetLength(1); j++)
                            {
                                var sheetName = sheetnames[0, j];
                                ExcelReference sheetRef = (ExcelReference)XlCall.Excel(XlCall.xlSheetId, sheetName);

                                ids.Add(sheetRef.SheetId.ToString());
                            }
                        }
                        catch (Exception ex)
                        {
                            logger.Error(ex);
                            ids = null;
                        }

                    }).Wait();
                }
                catch (Exception ex)
                {
                    logger.Error(ex);
                    ids = null;
                }
                return ids;
            });
        }
        #endregion
    }
}
