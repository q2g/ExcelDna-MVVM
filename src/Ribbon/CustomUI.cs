namespace ExcelDna_MVVM
{
    #region Usings
    using ExcelDna.Integration;
    using ExcelDna.Integration.CustomUI;
    using ExcelDna_MVVM.MVVM;
    using ExcelDna_MVVM.Ribbon;
    using ExcelDna_MVVM.Utils;
    using NLog;
    using System;
    using System.Collections.Generic;
    using System.Collections.Specialized;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Windows.Input;
    #endregion

    [ComVisible(true)]
    public class CustomUI : ExcelRibbon
    {
        #region Logger
        private static Logger logger = LogManager.GetCurrentClassLogger();
        #endregion

        #region CTOR
        public CustomUI()
        {
        }
        #endregion

        #region Variables & Properties
        private IRibbonUI ribbonUI;
        private List<BindingInfo> bindingInfos;
        IAddInInformation extRibbonData;
        private List<BoundControl> boundVMControls = new List<BoundControl>();

        internal static int? AppHwnd
        {
            get
            {
                try
                {
                    return (ExcelDnaUtil.Application as dynamic)?.Hwnd;
                }
                catch (Exception ex)
                {
                    logger.Error(ex);
                    return null;
                }
            }
        }
        #endregion

        #region Overrides
        public override string GetCustomUI(string RibbonID)
        {
            try
            {
                extRibbonData = FindRibbonDataImplementation();
                if (extRibbonData != null)
                {
                    extRibbonData.InvalidateRibbonCommand = new RelayCommand((o) =>
                    {
                        string param = null;
                        if (o is string strValue)
                            param = strValue;
                        InvalidateRibbon(param);

                    });
                    var ribbondefinition = RibbonDefinitionParser.ParseDefinition(extRibbonData.GetRibbonXML(), extRibbonData);
                    bindingInfos = ribbondefinition.Item2;

                    MVVMStatic.Adapter.VMCreated += Adapter_VMCreated;
                    MVVMStatic.Adapter.VMDeleted += Adapter_VMDeleted;

                    var createdVms = MVVMStatic.Adapter.AllVms;
                    foreach (var vms in createdVms)
                    {
                        foreach (var vm in vms.Value)
                        {
                            Adapter_VMCreated(MVVMStatic.Adapter, new VMEventArgs() { VM = vm, HWND = vms.Key });
                        }

                    }
                    return ribbondefinition.Item1;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return "";
        }
        #endregion

        #region RibbonHandler Functions
        public void OnAction(IRibbonControl control)
        {
            try
            {
                var ctrls = FindBoundControls(control.Id, control.GetHwnd(), RibbonBindingType.Command);
                foreach (var ctrl in ctrls)
                {
                    if (ctrl.Binding.CachedData is ICommand command)
                    {
                        command.Execute(null);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }

        public void OnActionToggle(IRibbonControl control, bool pressed)
        {
            try
            {
                OnAction(control);

                var toggles = FindBoundControls(control.Id, control.GetHwnd(), RibbonBindingType.TogglePressed);
                foreach (var toggle in toggles)
                {
                    if (toggle.Binding.CachedData is bool)//TODO: Remove this Workaround for setting a Value to a Binding which sourceObject don't have this Property.   
                    {
                        toggle.Binding.Value = pressed;
                    }
                }

            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }

        public bool GetPressed(IRibbonControl control)
        {
            try
            {
                return GetBindingValue<bool>(control, RibbonBindingType.TogglePressed);
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return false;
        }

        public bool GetEnabled(IRibbonControl control)
        {
            try
            {
                return GetBindingValue<bool>(control, RibbonBindingType.Enabled);
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return false;
        }

        public bool GetVisible(IRibbonControl control)
        {
            try
            {
                return GetBindingValue<bool>(control, RibbonBindingType.Visible);
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return false;
        }

        public string GetLabel(IRibbonControl control)
        {
            try
            {
                var bindinginfo = GetResourceLabelBinding(control.Id);
                if (bindinginfo != null)
                {
                    if (extRibbonData != null)
                    {
                        return extRibbonData.GetLocalizedString(bindinginfo.ResourceKey);
                    }
                    return bindinginfo.ResourceKey;
                }
                else
                {
                    return GetBindingValue<string>(control, RibbonBindingType.LabelBinding);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return "";
        }

        public void OnLoad(IRibbonUI ribbon)
        {
            try
            {
                logger.Info("onLoad: " + (ribbon != null).ToString());
                ribbonUI = ribbon;

                if (ribbonUI != null)
                {
                    ribbonUI.Invalidate();
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }

        public void OnItemsAction(IRibbonControl control, string id, int index)
        {
            try
            {
                var boundCollections = FindBoundControls(control.Id, control.GetHwnd(), RibbonBindingType.GalleryItemsSource);
                foreach (var boundCollectionControl in boundCollections)
                {
                    if (boundCollectionControl.Binding.CachedData is System.Collections.IList list)
                    {
                        if (index < list.Count)
                        {
                            var commandParameter = list[index];
                            var boundCommands = FindBoundControls(control.Id, control.GetHwnd(), RibbonBindingType.Command, boundCollectionControl.Binding.SourceObject);
                            var boundCommand = boundCommands.FirstOrDefault();
                            if (boundCommand != null)
                            {
                                if (boundCommand.Binding.CachedData is ICommand command)
                                {
                                    command?.Execute(commandParameter);
                                }
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

        public int GetItemCount(IRibbonControl control)
        {
            try
            {
                var boundItemsSources = FindBoundControls(control.Id, control.GetHwnd(), RibbonBindingType.GalleryItemsSource);

                foreach (var boundItemsSource in boundItemsSources)
                {
                    if (boundItemsSource.Binding.CachedData is System.Collections.IList items)
                    {
                        return items.Count;
                    }
                }

            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return 0;
        }

        public string GetItemID(IRibbonControl control, int index)
        {
            try
            {
                return GetItemBinding<string>(control, index, RibbonBindingType.ItemId);
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return "";
        }

        public string GetItemLabel(IRibbonControl control, int index)
        {
            try
            {
                return GetItemBinding<string>(control, index, RibbonBindingType.ItemLabel);
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return "";
        }

        public System.Drawing.Image GetItemImage(IRibbonControl control, int index)
        {
            try
            {
                return GetItemBinding<System.Drawing.Image>(control, index, RibbonBindingType.ItemImage);
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return null;
        }
        #endregion

        #region private Functions
        private BindingInfo GetResourceLabelBinding(string id)
        {
            return bindingInfos.FirstOrDefault(bind => bind.RibbonBindingType == RibbonBindingType.LabelFromResource
                                                             && bind.ID == id);

        }

        private T GetBindingValue<T>(IRibbonControl control, RibbonBindingType bindingType)
        {
            try
            {
                var ctrls = FindBoundControls(control.Id, control.GetHwnd(), bindingType);
                foreach (var ctrl in ctrls)
                {
                    if (ctrl.Binding.CachedData is T value)
                        return value;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return default(T);
        }

        private T GetItemBinding<T>(IRibbonControl control, int index, RibbonBindingType bindingType)
        {
            try
            {
                var ctrls = GetBoundItems(control, index, bindingType);
                foreach (var ctrl in ctrls)
                {
                    if (ctrl.Binding.CachedData is T value)
                        return value;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return default(T);
        }

        private List<BoundControl> GetBoundItems(IRibbonControl control, int index, RibbonBindingType bindingType)
        {
            List<BoundControl> items = new List<BoundControl>();
            try
            {
                object itemToFind = null;
                var boundLists = FindBoundControls(control.Id, control.GetHwnd(), RibbonBindingType.GalleryItemsSource);
                foreach (var boundList in boundLists)
                {
                    if (boundList.Binding.CachedData is System.Collections.IList list)
                    {
                        itemToFind = list[index];
                    }
                }
                if (itemToFind != null)
                {
                    foreach (var item in boundVMControls)
                    {
                        if (item.Binding.SourceObject == itemToFind && item.BindingInfo.RibbonBindingType == bindingType)
                        {
                            items.Add(item);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return items;
        }

        private List<BoundControl> FindBoundControls(string id, int? hwnd, RibbonBindingType type, object sourceObject = null)
        {
            List<BoundControl> erg = new List<BoundControl>();
            try
            {
                if (!hwnd.HasValue)
                    hwnd = -1;

                foreach (var boundControl in boundVMControls)
                {
                    if (boundControl.BindingInfo.ID == id
                        && (boundControl.Hwnd == hwnd || boundControl.Hwnd == -1)
                        && boundControl.BindingInfo.RibbonBindingType == type
                        && (sourceObject == null || boundControl.Binding.SourceObject == sourceObject))
                    {
                        erg.Add(boundControl);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return erg;
        }

        private void Adapter_VMCreated(object sender, EventArgs e)
        {
            try
            {
                if (e is VMEventArgs vmEventArgs)
                {
                    foreach (var bindingInfo in bindingInfos)
                    {
                        CreateBoundObject(vmEventArgs.VM, bindingInfo, vmEventArgs.HWND, vmEventArgs.VM);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }

        private void CreateBoundObject(object bindingSource, BindingInfo bindingInfo, int hwnd, object belongsToVm = null)
        {
            try
            {
                object collection = null;
                BoundControl boundObject = null;
                switch (bindingInfo.RibbonBindingType)
                {
                    case RibbonBindingType.TogglePressed:
                    case RibbonBindingType.Visible:
                    case RibbonBindingType.Enabled:
                    case RibbonBindingType.LabelBinding:
                    case RibbonBindingType.ItemId:
                    case RibbonBindingType.ItemLabel:
                    case RibbonBindingType.ItemImage:
                    case RibbonBindingType.Invalidation:
                        boundObject = new BoundControl();
                        boundObject.Binding = new BindingObject(bindingSource, bindingInfo.BindingPath, (eArgs) =>
                        {
                            InvalidateRibbon(bindingInfo.ID);
                        }, true);
                        break;


                    case RibbonBindingType.Command:
                    case RibbonBindingType.ToggleCommand:
                        boundObject = new BoundControl();
                        boundObject.Binding = new BindingObject(bindingSource, bindingInfo.BindingPath, null, false);
                        break;

                    case RibbonBindingType.GalleryItemsSource:
                        boundObject = new BoundControl();
                        boundObject.Binding = new BindingObject(bindingSource, bindingInfo.BindingPath, (eArgs) =>
                        {
                            if (eArgs.NewValue is INotifyCollectionChanged collectionChangedNew)
                            {
                                collectionChangedNew.CollectionChanged += GalleryCollectionChanged;
                                collection = eArgs.NewValue;
                            }

                            if (eArgs.OldValue is INotifyCollectionChanged collectionChangedOld)
                            {
                                collectionChangedOld.CollectionChanged -= GalleryCollectionChanged;
                            }
                        }, false);
                        break;

                    default:
                        break;
                }
                if (boundObject != null)
                {
                    boundObject.BindingInfo = bindingInfo;
                    boundObject.Hwnd = hwnd;
                    boundObject.BelongsToVM = belongsToVm;
                    boundVMControls.Add(boundObject);
                    logger.Trace($"+++++++++++++++++Created BoundObject: {boundObject} +++++++++++++++Count of  BoundObjects:{boundVMControls.Count}");

                    if (bindingInfo.RibbonBindingType == RibbonBindingType.GalleryItemsSource)
                        GalleryCollectionChanged(collection, new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset));
                }


            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }

        private void Adapter_VMDeleted(object sender, EventArgs e)
        {
            try
            {
                if (e is VMEventArgs vmEventArgs)
                {
                    var list = boundVMControls.ToList();
                    foreach (var boundControl in list)
                    {
                        if (boundControl.BelongsToVM == vmEventArgs.VM)
                        {
                            boundVMControls.Remove(boundControl);
                            logger.Trace($"-----------------------Removed BoundControl: {boundControl}--------------Count of BoundObjects:{boundVMControls.Count}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }

        private BoundControl FindBoundCollectionControlsBySourceObject<T>(object source) where T : class
        {
            try
            {
                foreach (var ctrl in boundVMControls)
                {
                    if (ctrl.BindingInfo.RibbonBindingType == RibbonBindingType.GalleryItemsSource)
                    {
                        var obj = ctrl.Binding.CachedData as T;
                        if (obj is T && obj == source)
                            return ctrl;
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return null;
        }


        private void GalleryCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            switch (e.Action)
            {
                case NotifyCollectionChangedAction.Add:
                    var boundControl = FindBoundCollectionControlsBySourceObject<System.Collections.IList>(sender);
                    if (boundControl != null)
                    {
                        foreach (var subbindingInfo in boundControl.BindingInfo.SubInfos)
                        {
                            foreach (var item in e.NewItems)
                            {
                                CreateBoundObject(item, subbindingInfo, boundControl.Hwnd, boundControl.Binding.SourceObject);
                            }
                        }
                        InvalidateRibbon(boundControl.BindingInfo.ID);
                    }
                    break;
                case NotifyCollectionChangedAction.Remove:
                    boundControl = FindBoundCollectionControlsBySourceObject<System.Collections.IList>(sender);
                    if (boundControl != null)
                    {
                        List<BoundControl> toDelete = new List<BoundControl>();
                        foreach (var subbindingInfo in boundControl.BindingInfo.SubInfos)
                        {
                            foreach (var item in e.OldItems)
                            {
                                foreach (var control in boundVMControls)
                                {
                                    if (control.Binding.SourceObject == item)
                                    {
                                        toDelete.Add(control);
                                    }
                                }
                            }
                        }
                        foreach (var item in toDelete)
                        {
                            boundVMControls.Remove(item);
                            item.Dispose();
                        }
                        InvalidateRibbon(boundControl.BindingInfo.ID);
                    }
                    break;
                case NotifyCollectionChangedAction.Replace:
                    break;
                case NotifyCollectionChangedAction.Move:
                    break;
                case NotifyCollectionChangedAction.Reset:
                    boundControl = FindBoundCollectionControlsBySourceObject<System.Collections.IList>(sender);
                    if (boundControl != null)
                    {
                        if (sender is System.Collections.IEnumerable items)
                        {
                            foreach (var subbindingInfo in boundControl.BindingInfo.SubInfos)
                            {
                                foreach (var item in items)
                                {
                                    CreateBoundObject(item, subbindingInfo, boundControl.Hwnd, boundControl.Binding.SourceObject);
                                }
                            }
                        }
                        InvalidateRibbon(boundControl.BindingInfo.ID);
                    }

                    break;
                default:
                    break;
            }
            if (e.Action == NotifyCollectionChangedAction.Add)
            {

            }

            if (e.Action == NotifyCollectionChangedAction.Remove)
            {

            }

            if (e.Action == NotifyCollectionChangedAction.Reset)
            {

            }

        }

        private void InvalidateRibbon(string id = null)
        {
            try
            {
                if (ribbonUI != null)
                {
                    if (id != null)
                    {
                        ribbonUI.InvalidateControl(id);
                    }
                    else
                    {
                        ribbonUI.Invalidate();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }

        private IAddInInformation FindRibbonDataImplementation()
        {
            IAddInInformation retval = null;
            try
            {
                var types = TypeUtils.GetTypesImplementingInterface<IAddInInformation>();
                if (types.Count > 0)
                {
                    var typeToCreate = types.First();
                    retval = new AddInInformationWrapper(Activator.CreateInstance(typeToCreate));
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return retval;
        }
        #endregion
    }
}

