namespace ExcelDna_MVVM.MVVM
{
    #region Usings
    using ExcelDna.Integration;
    using ExcelDna_MVVM.Environment;
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
        private NetOffice.ExcelApi.Application Application;
        object sheetID2VMsLock = new object();
        private Dictionary<string, List<object>> sheetID2VMs = new Dictionary<string, List<object>>();

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
            Application = new Application(null, ExcelDnaUtil.Application);
            Application.NewWorkbookEvent += Application_NewWorkbookEvent;
            Application.WorkbookNewSheetEvent += Application_WorkbookNewSheetEvent;
            Application.WorkbookActivateEvent += Application_WorkbookActivateEvent;
            Application.SheetActivateEvent += Application_SheetActivateEvent;
            CreateVMsForApplication(Application as dynamic);
        }
        #endregion

        #region public Functions       
        #endregion

        #region Eventhandler
        private void Application_WorkbookActivateEvent(Workbook wb)
        {
            try
            {
                logger.Trace($"workbook activated {wb.Name}");
                RemoveUnusedVms();
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            wb.DisposeChildInstances();
        }

        private void Application_SheetActivateEvent(COMObject sheet)
        {
            logger.Trace($"sheet activated {(sheet as Worksheet).Name}");
            RemoveUnusedVms();
            sheet.DisposeChildInstances();
        }

        private void Application_NewWorkbookEvent(Workbook wb)
        {
            try
            {
                ConvertWorkbookAsync(wb).ContinueWith((res) =>
                {
                    if (!res.IsFaulted && res.Result != null)
                    {

                        foreach (var hwnd in res.Result.hwnds)
                        {
                            CreateVMImplementations<IWorkbookVM>(hwnd);
                            CreateSheetVMsFromWorkbookAsync(res.Result);
                        }

                    }
                });
                wb.DisposeChildInstances();
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
                ConvertWorkbookAsync(wb).ContinueWith((res) =>
                {
                    if (!res.IsFaulted && res != null)
                        CreateSheetVMsFromWorkbookAsync(res.Result);
                });
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            wb.DisposeChildInstances();
            sheet.DisposeChildInstances();
        }
        #endregion

        #region private Functions
        private Task RemoveUnusedVms()
        {
            Task retval = new Task(() => { });
            try
            {
                var app = Application;

                List<Task<WorkbookData>> tasks = new List<Task<WorkbookData>>();
                foreach (var wb in app.Workbooks)
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

                CreateVMImplementations<IAppVM>(-1);

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

        private List<object> CreateVMImplementations<T>(int hwnd) where T : IVM
        {
            List<object> createdVms = new List<object>();
            try
            {
                var types = vmImplementationTypes[typeof(T)];
                createdVms = types.Select((type) =>
                  {
                      try
                      {
                          logger.Info($"Create VM for Type: {type?.FullName}");
                          var vm = Activator.CreateInstance(type);
                          //if (typeof(T) == typeof(IAppVM))
                          //{
                          //    (vm as IAppVM).WindowService = new WindowService()
                          //    {
                          //        RibbonHeight = Application.CommandBars["Ribbon"].Height,
                          //        RibbonWidth = Application.CommandBars["Ribbon"].Width
                          //    };
                          //}
                          //var piWindowService = type.GetProperty("WindowService");
                          //if (piWindowService != null)
                          //{
                          //    if (piWindowService.PropertyType.FullName == typeof(WindowService).FullName)
                          //    {
                          //        piWindowService.SetValue(vm, new WindowService()
                          //        {
                          //            RibbonHeight = Application.CommandBars["Ribbon"].Height,
                          //            RibbonWidth = Application.CommandBars["Ribbon"].Width
                          //        });
                          //    }
                          //}
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
                                                if (!vms.ContainsKey(hwnd))
                                                    vms.Add(hwnd, new List<object>());
                                                vms[hwnd].Add(vm);
                                                logger.Trace(() => GetVMsCount());
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

        private Task<WorkbookData> ConvertWorkbookAsync(Workbook wb)
        {
            WorkbookData wbd = new WorkbookData();
            wbd.Name = wb.Name;
            wbd.hwnds = wb.Windows.Select(win => win.Hwnd).ToList();
            return GetSheetIdsFromWorkbookAsync(wbd).ContinueWith((res) =>
            {
                if (!res.IsFaulted && res.Result != null)
                    wbd.sheetIds.AddRange(res.Result);
                return wbd;
            });


        }

        private Task CreateSheetVMsFromWorkbookAsync(WorkbookData wbd)
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
                                        var impls = CreateVMImplementations<IWorksheetVM>(hwnd);
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
