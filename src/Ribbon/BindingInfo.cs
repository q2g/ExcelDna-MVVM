namespace ExcelDna_MVVM.Ribbon
{
    #region Usings
    using System.Collections.Generic;
    #endregion

    public enum RibbonBindingType
    {
        Invalidation,
        Command,
        ToggleCommand,
        TogglePressed,
        Enabled,
        Visible,
        GalleryItemsSource,
        ItemId,
        ItemLabel,
        ItemImage,
        LabelBinding,
        LabelFromResource
    }



    public class BindingInfo
    {
        public string BindingPath { get; set; }
        public RibbonBindingType RibbonBindingType { get; set; }
        public string ID { get; set; }
        public string ParentID { get; set; }
        public string ResourceKey { get; set; }
        public List<BindingInfo> SubInfos { get; set; } = new List<BindingInfo>();
    }
}
