namespace ExcelDna_MVVM.MVVM
{
    #region Usings
    using ExcelDna.Integration;
    using ExcelDna_MVVM.Environment;
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
        private object vmsLock = new object();
        private Dictionary<int, List<IVM>> vms = new Dictionary<int, List<IVM>>();
        private Dictionary<Type, PropertyInfo> propertyInfos = new Dictionary<Type, PropertyInfo>();
        private Dictionary<Type, List<Type>> vmImplementationTypes = new Dictionary<Type, List<Type>>();

        public Dictionary<int, List<IVM>> AllVms //TODO:
        {
            get
            {
                return vms;
            }
        }
        private NetOffice.ExcelApi.Application Application;
        private Dictionary<string, List<IVM>> sheetID2VMs;
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
            Application.SheetBeforeDeleteEvent += Application_SheetBeforeDeleteEvent;
            CreateVMsForApplication(Application as dynamic);
        }
        #endregion

        #region public Functions       
        #endregion

        #region private Functions
        private void CreateVMsForApplication(dynamic app)
        {
            try
            {
                if (vms == null)
                    vms = new Dictionary<int, List<IVM>>();

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

        private List<IVM> CreateVMImplementations<T>(int hwnd) where T : IVM
        {
            List<IVM> createdVms = new List<IVM>();
            try
            {
                var types = vmImplementationTypes[typeof(T)];
                foreach (var type in types)
                {
                    try
                    {
                        logger.Info($"Create VM for Type: {type?.FullName}");
                        var vm = (T)Activator.CreateInstance(type);
                        createdVms.Add(vm);

                        if (!vms.ContainsKey(hwnd))
                            vms.Add(hwnd, new List<IVM>());
                        vms[hwnd].Add(vm);

                        logger.Trace(() => GetVMsCount());
                        VMCreated?.Invoke(this, new VMEventArgs() { VM = vm, HWND = hwnd });
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return createdVms;
        }

        private string GetVMsCount()
        {
            var allVms = vms.SelectMany(ele => ele.Value).ToList();
            return $"------------------------------------AppVMs:{allVms.Where(vm => vm is IAppVM).Count()}, WorkbookVMs:{allVms.Where(vm => vm is IWorkbookVM).Count()}, WorksheetVMs:{allVms.Where(vm => vm is IWorksheetVM).Count()}";
        }

        private void Application_NewWorkbookEvent(Workbook wb)
        {
            Task.Run(() =>
            {
                try
                {
                    int hwnd = GetHwndFromWorkbook(wb);
                    if (hwnd != -2)
                    {
                        CreateVMImplementations<IWorkbookVM>(hwnd);
                        CreateSheetVMsFromWorkbookAsync(wb);
                    }
                }
                catch (Exception ex)
                {
                    logger.Error(ex);
                }
                wb.DisposeChildInstances();
            });
        }

        private void Application_WorkbookNewSheetEvent(Workbook wb, COMObject sheet)
        {
            try
            {
                CreateSheetVMsFromWorkbookAsync(wb);
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            wb.DisposeChildInstances();
            sheet.DisposeChildInstances();
        }

        private void Application_SheetBeforeDeleteEvent(COMObject Sh)
        {

        }

        private Task CreateSheetVMsFromWorkbookAsync(Workbook wb)
        {
            var ids = new List<string>();
            return EnvironmentAdapter.QueueAction(() =>
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
                }
            })
            .ContinueWith((res) =>
            {
                foreach (var id in ids)
                {
                    if (!sheetID2VMs.ContainsKey(id))
                        sheetID2VMs.Add(id, CreateVMImplementations<IWorksheetVM>(wb.Windows[0].Hwnd));
                }
            });
        }

        private int GetHwndFromWorkbook(NetOffice.ExcelApi.Workbook wb)
        {
            if (wb.Windows.Count > 0)//TODO: How can there be Multiple Hwnd's for one Workbook?
            {
                return wb.Windows[1].Hwnd;
            }
            else
            {
                logger.Warn($"New Workbook {wb?.Name ?? "unknown"} does not have a window");
            }
            return -2;
        }
        #endregion
    }
}
