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

            string result = MVVMStatic.RibbonXml; //TODO: Find a Way to provide the xml
            var ribbondefinition = RibbonDefinitionParser.ParseDefinition(result);
            bindingInfos = ribbondefinition.Item2;

            MVVMStatic.Adapter.VMCreated += Adapter_VMCreated;

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
        #endregion

        #region RibbonHandler Functions
        public void OnAction(IRibbonControl control)
        {
            try
            {
                var ctrls = FindBoundControls(control.Id, control.GetHwnd(), RibbonBindingType.Command);
                foreach (var ctrl in ctrls)
                {
                    if (ctrl.Binding.Value is ICommand command)
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
                    if (toggle.Binding.Value is bool)//TODO: Remove this Workaround for setting a Value to a Binding which sourceObject don't have this Property.   
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
                return GetBoolBinding(control, RibbonBindingType.TogglePressed);
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return false;
        }

        public bool getEnabled(IRibbonControl control)
        {
            try
            {
                return GetBoolBinding(control, RibbonBindingType.Enabled);
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return false;
        }

        public bool getVisible(IRibbonControl control)
        {
            try
            {
                return GetBoolBinding(control, RibbonBindingType.Visible);
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return false;
        }

        public void onLoad(IRibbonUI ribbon)
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
                    if (boundCollectionControl.Binding.Value is System.Collections.IList list)
                    {
                        if (index < list.Count)
                        {
                            var commandParameter = list[index];
                            var boundCommands = FindBoundControls(control.Id, control.GetHwnd(), RibbonBindingType.Command, boundCollectionControl.Binding.SourceObject);
                            var boundCommand = boundCommands.FirstOrDefault();
                            if (boundCommand != null)
                            {
                                if (boundCommand.Binding.Value is ICommand command)
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
                    if (boundItemsSource.Binding.Value is System.Collections.IList items)
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
        private List<BoundControl> GetBoundItems(IRibbonControl control, int index, RibbonBindingType bindingType)
        {
            List<BoundControl> items = new List<BoundControl>();
            try
            {
                object itemToFind = null;
                var boundLists = FindBoundControls(control.Id, control.GetHwnd(), RibbonBindingType.GalleryItemsSource);
                foreach (var boundList in boundLists)
                {
                    if (boundList.Binding.Value is System.Collections.IList list)
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

        private bool GetBoolBinding(IRibbonControl control, RibbonBindingType bindingType)
        {
            try
            {
                var ctrls = FindBoundControls(control.Id, control.GetHwnd(), bindingType);
                foreach (var ctrl in ctrls)
                {
                    if (ctrl.Binding.Value is bool boolvalue)
                        return boolvalue;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return false;
        }

        private T GetItemBinding<T>(IRibbonControl control, int index, RibbonBindingType bindingType)
        {
            try
            {
                var ctrls = GetBoundItems(control, index, bindingType);
                foreach (var ctrl in ctrls)
                {
                    if (ctrl.Binding.Value is T value)
                        return value;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return default(T);
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
                        BoundControl boundObject = new BoundControl();
                        switch (bindingInfo.RibbonBindingType)
                        {
                            case RibbonBindingType.TogglePressed:
                            case RibbonBindingType.Visible:
                            case RibbonBindingType.Enabled:
                            case RibbonBindingType.Invalidation:
                                boundObject.Binding = new BindingObject(vmEventArgs.VM, bindingInfo.BindingPath, (eArgs) =>
                                {
                                    InvalidateRibbon(bindingInfo.ID);
                                }, true);
                                break;


                            case RibbonBindingType.Command:
                            case RibbonBindingType.ToggleCommand:
                                boundObject.Binding = new BindingObject(vmEventArgs.VM, bindingInfo.BindingPath, null, false);
                                break;

                            case RibbonBindingType.GalleryItemsSource:
                                boundObject.Binding = new BindingObject(vmEventArgs.VM, bindingInfo.BindingPath, (eArgs) =>
                                {
                                    var handler = new NotifyCollectionChangedEventHandler((s, eventargs) =>
                                    {
                                        var bi = bindingInfo;
                                        GalleryCollectionChanged(s, eventargs, bi, vmEventArgs.HWND);
                                    });

                                    if (eArgs.NewValue is INotifyCollectionChanged collectionChangedNew)
                                    {
                                        collectionChangedNew.CollectionChanged += handler;
                                        handler(collectionChangedNew, new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset));
                                    }

                                    if (eArgs.OldValue is INotifyCollectionChanged collectionChangedOld)
                                    {
                                        collectionChangedOld.CollectionChanged -= handler;
                                    }
                                }, false);
                                break;

                            default:
                                break;
                        }
                        if (boundObject != null)
                        {
                            boundObject.BindingInfo = bindingInfo;
                            boundObject.Hwnd = vmEventArgs.HWND;
                            boundVMControls.Add(boundObject);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }

        private void GalleryCollectionChanged(object sender, NotifyCollectionChangedEventArgs e, BindingInfo bindingInfo, int hwnd)
        {
            if (e.Action == NotifyCollectionChangedAction.Add)
            {
                foreach (var subbindingInfo in bindingInfo.SubInfos)
                {
                    foreach (var item in e.NewItems)
                    {
                        BoundControl boundItem = new BoundControl
                        {
                            Binding = new BindingObject(item, subbindingInfo.BindingPath, (eArgs) =>
                            {
                                InvalidateRibbon(bindingInfo.ID);
                            }, true),
                            BindingInfo = subbindingInfo,
                            Hwnd = hwnd
                        };
                        boundVMControls.Add(boundItem);
                    }
                }
            }
            if (e.Action == NotifyCollectionChangedAction.Remove)
            {
                List<BoundControl> toDelete = new List<BoundControl>();
                foreach (var subbindingInfo in bindingInfo.SubInfos)
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
            }
            if (e.Action == NotifyCollectionChangedAction.Reset)
            {
                if (sender is System.Collections.IEnumerable items)
                {
                    foreach (var subbindingInfo in bindingInfo.SubInfos)
                    {
                        foreach (var item in items)
                        {
                            BoundControl boundItem = new BoundControl
                            {
                                Binding = new BindingObject(item, subbindingInfo.BindingPath, (eArgs) =>
                                {
                                    InvalidateRibbon(subbindingInfo.ParentID);
                                }, true),
                                BindingInfo = subbindingInfo,
                                Hwnd = hwnd
                            };
                            boundVMControls.Add(boundItem);
                        }
                    }
                }
            }
            InvalidateRibbon(bindingInfo.ID);
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
        #endregion
    }
}

