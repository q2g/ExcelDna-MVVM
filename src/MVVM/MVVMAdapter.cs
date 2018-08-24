namespace ExcelDna_MVVM.MVVM
{
    #region Usings
    using ExcelDna.Integration;
    using ExcelDna_MVVM.Utils;
    using NLog;
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Reflection;
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

        public Dictionary<int, List<IVM>> AllVms
        {
            get
            {
                return vms;
            }
        }
        #endregion

        #region Events
        public event EventHandler VMCreated;
        #endregion

        #region IExcelAddIn
        public void AutoClose()
        {
        }

        public void AutoOpen()
        {
            MVVMStatic.Adapter = this;
            CreateVMsForApplication(ExcelDnaUtil.Application as dynamic);
        }
        #endregion

        #region public Functions       
        #endregion

        #region Hwnd / VM Helper Functions 
        private void CreateVMsForApplication(dynamic app)
        {
            try
            {
                if (vms == null)
                    vms = new Dictionary<int, List<IVM>>();

                if (!vms.ContainsKey(-1))
                    vms.Add(-1, new List<IVM>());
                var appVms = GetVMImplementations<IAppVM>(-1);
                vms[-1].AddRange(appVms);


                foreach (var workbook in app.Workbooks)
                {
                    int hwnd = workbook.Windows[0].Hwnd;
                    if (!vms.ContainsKey(hwnd))
                        vms.Add(hwnd, new List<IVM>());

                    var workbookVms = GetVMImplementations<IWorkbookVM>(hwnd);
                    vms[hwnd].AddRange(workbookVms);

                    foreach (var item in workbook.Worksheets)
                    {
                        var sheetVms = GetVMImplementations<IWorksheetVM>(hwnd);
                        vms[hwnd].AddRange(sheetVms);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }

        private List<T> GetVMImplementations<T>(int hwnd) where T : IVM
        {
            List<T> vminstances = new List<T>();
            try
            {

                var types = TypeUtils.GetTypesImplementingInterface<IAppVM>();
                foreach (var type in types)
                {
                    try
                    {
                        logger.Info($"Create VM for Type: {type}");
                        var vm = (T)Activator.CreateInstance(type);
                        vminstances.Add(vm);
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
            return vminstances;
        }
        #endregion

        #region private Functions
        private static void VM_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
        }
        #endregion
    }
}
