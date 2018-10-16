namespace ExcelDna_MVVM
{
    #region Usings
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
    using System.Threading.Tasks;
    using System.Windows.Input;
    using System.Windows.Threading;
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
        private IAddInInformation addInInformation;
        private object boundVMControlsLock = new object();
        private List<BoundControl> boundVMControls = new List<BoundControl>();
        private Dispatcher currentdispatcher;
        #endregion

        #region Overrides
        public override string GetCustomUI(string RibbonID)
        {
            currentdispatcher = Dispatcher.CurrentDispatcher;
            try
            {
                addInInformation = FindRibbonDataImplementation();
                if (addInInformation != null)
                {
                    addInInformation.InvalidateRibbonCommand = new RelayCommand((o) =>
                    {
                        string param = null;
                        if (o is string strValue)
                            param = strValue;
                        InvalidateRibbon(param);

                    });
                    var ribbondefinition = RibbonDefinitionParser.ParseDefinition(addInInformation.GetRibbonXML(), addInInformation);
                    bindingInfos = ribbondefinition.Item2;

                    MVVMStatic.Adapter.VMCreated += Adapter_VMCreated;
                    MVVMStatic.Adapter.VMDeleted += Adapter_VMDeleted;

                    var createdVms = MVVMStatic.Adapter.AllVms;
                    foreach (var item in createdVms)
                    {
                        Adapter_VMCreated(MVVMStatic.Adapter, new VMEventArgs() { VMs = item.Value, HWND = item.Key });
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
                var bindinginfo = GetResourceBinding(control.Id, RibbonBindingType.LabelFromResource);
                if (bindinginfo != null)
                {
                    if (addInInformation != null)
                    {
                        return (string)addInInformation.GetResource(bindinginfo.ResourceKey);
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

        public System.Drawing.Image GetImage(IRibbonControl control)
        {
            try
            {
                //return LocalizeDictionary.Instance.GetLocalizedObject(ribbon.Tag + "_Image", null, LocalizeDictionary.Instance.Culture) as System.Drawing.Bitmap;
                var bindinginfo = GetResourceBinding(control.Id, RibbonBindingType.ImageFromResource);
                if (bindinginfo != null)
                {
                    if (addInInformation != null)
                    {
                        return (System.Drawing.Image)addInInformation.GetResource(bindinginfo.ResourceKey + "_Image");
                    }
                    return null;
                }
                else
                {
                    return GetBindingValue<System.Drawing.Image>(control, RibbonBindingType.ImageBinding);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
                return null;
            }
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

        public void OnSelectedChanged(IRibbonControl control, string text)
        {
            try
            {
                var selchanged = FindBoundControls(control.Id, control.GetHwnd(), RibbonBindingType.ComboboxSelectedChanged);
                foreach (var ctrl in selchanged)
                {
                    if (ctrl.Binding.CachedData is string)//TODO: Remove this Workaround for setting a Value to a Binding for sourceObject which don't have this Property.   
                    {
                        if (text == null)
                            text = string.Empty; //Never set BindingValue to null, or the Binding is dead //TODO: find a Solution for this
                        ctrl.Binding.Value = text;
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }

        }

        public string GetText(IRibbonControl control)
        {
            try
            {
                return GetBindingValue<string>(control, RibbonBindingType.ComboboxSelectedText);
                //TODO: How to Detect when a string Binding is bound but string is null???
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return null;
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
                            var boundCommands = FindBoundControls(control.Id, control.GetHwnd(), RibbonBindingType.Command);//boundCollectionControl.Binding.SourceObject
                            foreach (var boundCommand in boundCommands)
                            {
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
        private BindingInfo GetResourceBinding(string id, RibbonBindingType bindingType)
        {
            return bindingInfos.FirstOrDefault(bind => bind.RibbonBindingType == bindingType
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
                    List<BoundControl> boundcontrols = new List<BoundControl>();
                    Task.Run(() =>
                    {
                        foreach (var vm in vmEventArgs.VMs)
                        {
                            foreach (var bindingInfo in bindingInfos)
                            {
                                var ctrl = CreateBoundObject(vm, bindingInfo, vmEventArgs.HWND, vm);
                                if (ctrl != null)
                                {
                                    boundcontrols.Add(ctrl);
                                }
                            }
                        }
                    }).ContinueWith((task) =>
                    {
                        lock (boundVMControlsLock)
                        {
                            boundVMControls.AddRange(boundcontrols);
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }

        private BoundControl CreateBoundObject(object bindingSource, BindingInfo bindingInfo, int hwnd, object belongsToVm = null)
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
                    case RibbonBindingType.ComboboxSelectedChanged:
                    case RibbonBindingType.ComboboxSelectedText:
                    case RibbonBindingType.Invalidation:
                        boundObject = new BoundControl();
                        currentdispatcher.Invoke(() =>
                        {
                            boundObject.Binding = new BindingObject(bindingSource, bindingInfo.BindingPath, (eArgs) =>
                            {
                                InvalidateRibbon(bindingInfo.ID);
                            }, true);
                        });
                        break;
                    case RibbonBindingType.Command:
                    case RibbonBindingType.ToggleCommand:
                        boundObject = new BoundControl();
                        currentdispatcher.Invoke(() =>
                        {
                            boundObject.Binding = new BindingObject(bindingSource, bindingInfo.BindingPath, null, false);
                        });
                        break;

                    case RibbonBindingType.GalleryItemsSource:
                        boundObject = new BoundControl();
                        boundObject.BindingInfo = bindingInfo;
                        boundObject.Hwnd = hwnd;
                        boundObject.BelongsToVM = belongsToVm;
                        currentdispatcher.Invoke(() =>
                        {
                            boundObject.Binding = new BindingObject(bindingSource, bindingInfo.BindingPath, null, false);
                        });
                        boundObject.Binding.OnChanged = (eArgs) =>
                         {
                             if (eArgs.NewValue is INotifyCollectionChanged collectionChangedNew)
                             {
                                 collectionChangedNew.CollectionChanged += GalleryCollectionChanged;
                                 collection = eArgs.NewValue;
                                 GalleryCollectionChanged(collection, new BoundControlCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset) { Control = boundObject });
                             }

                             if (eArgs.OldValue is INotifyCollectionChanged collectionChangedOld)
                             {
                                 if (eArgs.OldValue is System.Collections.IEnumerable items)
                                 {
                                     var list = FindBoundControlsByCollectionItems(items);
                                     RemoveBoundControls(list);
                                 }
                                 collectionChangedOld.CollectionChanged -= GalleryCollectionChanged;
                             }
                         };
                        break;

                    default:
                        break;
                }
                if (boundObject != null)
                {
                    boundObject.BindingInfo = bindingInfo;
                    boundObject.Hwnd = hwnd;
                    boundObject.BelongsToVM = belongsToVm;
                    logger.Trace($"+++++++++++++++++Created BoundObject: {boundObject} +++++++++++++++Count of  BoundObjects:{boundVMControls.Count}");
                }
                return boundObject;

            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return null;
        }

        private List<BoundControl> FindBoundControlsByCollectionItems(System.Collections.IEnumerable source)
        {
            List<BoundControl> retval = new List<BoundControl>();
            foreach (var item in source)
            {
                retval.AddRange(boundVMControls.Where(ctrl => ctrl != null && ctrl.Binding != null && ctrl.Binding.SourceObject == item).ToList());
            }
            return retval;
        }

        private void Adapter_VMDeleted(object sender, EventArgs e)
        {
            try
            {
                if (e is VMEventArgs vmEventArgs)
                {
                    List<BoundControl> toDelete = new List<BoundControl>();
                    foreach (var boundControl in boundVMControls)
                    {
                        foreach (var vm in vmEventArgs.VMs)
                        {
                            if (boundControl.BelongsToVM == vm)
                            {
                                toDelete.Add(boundControl);
                            }
                        }
                    }
                    RemoveBoundControls(toDelete);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }

        private void RemoveBoundControls(IList<BoundControl> list)
        {
            lock (boundVMControlsLock)
            {
                foreach (var boundControl in list)
                {
                    boundVMControls.Remove(boundControl);
                    logger.Trace($"-----------------------Removed BoundControl: {boundControl}--------------Count of BoundObjects:{boundVMControls.Count}");
                }
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
                        List<BoundControl> toAdd = new List<BoundControl>();
                        foreach (var subbindingInfo in boundControl.BindingInfo.SubInfos)
                        {
                            foreach (var item in e.NewItems)
                            {
                                toAdd.Add(CreateBoundObject(item, subbindingInfo, boundControl.Hwnd, boundControl.Binding.SourceObject));
                            }
                        }
                        lock (boundVMControlsLock)
                        {
                            boundVMControls.AddRange(toAdd);
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
                                foreach (var ctrl in boundVMControls)
                                {
                                    if (ctrl.Binding.SourceObject == item)
                                    {
                                        toDelete.Add(ctrl);
                                    }
                                }
                            }
                        }
                        lock (boundVMControlsLock)
                        {
                            foreach (var item in toDelete)
                            {
                                boundVMControls.Remove(item);
                                item.Dispose();
                            }
                        }
                        InvalidateRibbon(boundControl.BindingInfo.ID);
                    }
                    break;
                case NotifyCollectionChangedAction.Replace:
                    break;
                case NotifyCollectionChangedAction.Move:
                    break;
                case NotifyCollectionChangedAction.Reset:
                    BoundControl control;
                    if (e is BoundControlCollectionChangedEventArgs controlArgs)
                    {
                        control = controlArgs.Control;
                    }
                    else
                    {
                        control = FindBoundCollectionControlsBySourceObject<System.Collections.IList>(sender);
                    }
                    if (control != null)
                    {

                        if (sender is System.Collections.IEnumerable items)
                        {
                            List<BoundControl> toAdd = new List<BoundControl>();
                            foreach (var subbindingInfo in control.BindingInfo.SubInfos)
                            {
                                foreach (var item in items)
                                {
                                    toAdd.Add(CreateBoundObject(item, subbindingInfo, control.Hwnd, control.Binding.SourceObject));
                                }
                            }
                            lock (boundVMControlsLock)
                            {
                                boundVMControls.AddRange(toAdd);
                            }
                        }
                        InvalidateRibbon(control.BindingInfo.ID);
                    }

                    break;
                default:
                    break;
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
                    retval = new AddInInformationProxy(Activator.CreateInstance(typeToCreate));
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

