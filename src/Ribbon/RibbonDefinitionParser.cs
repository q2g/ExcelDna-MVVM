namespace ExcelDna_MVVM.Ribbon
{
    using NLog;
    #region Usings
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Xml.Linq;
    #endregion

    public class RibbonDefinitionParser
    {
        #region Logger
        private static Logger logger = LogManager.GetCurrentClassLogger();
        #endregion

        #region static Functions        
        public static Tuple<string, List<BindingInfo>> ParseDefinition(string ribbonDefinitionText, IExtendedRibbonData extRibbondata)
        {
            try
            {
                string localizationPrefix = "";
                var bindingInfos = new List<BindingInfo>();
                XElement root = XElement.Parse(ribbonDefinitionText);

                var attr = root.Attributes().FirstOrDefault(atr => atr.Name.LocalName == "localizationPrefix");
                if (attr != null)
                {
                    localizationPrefix = attr.Value + ":";
                    attr.Remove();
                }
                var ribbonelement = root.Descendants().Where(ele => ele.Name.LocalName == "ribbon")
                    .FirstOrDefault();
                if (ribbonelement != null)//gallery
                {


                    bindingInfos.AddRange(ReplaceBindingAndExtractBindingInfo(ribbonelement, "button", "onAction", nameof(CustomUI.OnAction), RibbonBindingType.Command));
                    bindingInfos.AddRange(ReplaceBindingAndExtractBindingInfo(ribbonelement, "toggleButton", "getPressed", nameof(CustomUI.GetPressed), RibbonBindingType.TogglePressed));
                    bindingInfos.AddRange(ReplaceBindingAndExtractBindingInfo(ribbonelement, "toggleButton", "onAction", nameof(CustomUI.OnActionToggle), RibbonBindingType.Command));
                    bindingInfos.AddRange(ReplaceBindingAndExtractBindingInfo(ribbonelement, "", "getEnabled", nameof(CustomUI.GetEnabled), RibbonBindingType.Enabled));
                    bindingInfos.AddRange(ReplaceBindingAndExtractBindingInfo(ribbonelement, "", "getVisible", nameof(CustomUI.GetVisible), RibbonBindingType.Visible));
                    bindingInfos.AddRange(ReplaceBindingAndExtractBindingInfo(ribbonelement, "gallery", "itemssource", "#remove", RibbonBindingType.GalleryItemsSource));
                    bindingInfos.AddRange(ReplaceBindingAndExtractBindingInfo(ribbonelement, "gallery", "onAction", nameof(CustomUI.OnItemsAction), RibbonBindingType.Command));
                    bindingInfos.AddRange(ReplaceBindingAndExtractBindingInfo(ribbonelement, "gallery", "getItemCount", nameof(CustomUI.GetItemCount), RibbonBindingType.Invalidation));
                    bindingInfos.AddRange(ReplaceBindingAndExtractBindingInfo(ribbonelement, "button", "getLabel", nameof(CustomUI.GetLabel), RibbonBindingType.LabelBinding, null, extRibbondata, localizationPrefix));
                    bindingInfos.AddRange(ReplaceBindingAndExtractBindingInfo(ribbonelement, "toggleButton", "getLabel", nameof(CustomUI.GetLabel), RibbonBindingType.LabelBinding, null, extRibbondata, localizationPrefix));
                    bindingInfos.AddRange(ReplaceBindingAndExtractBindingInfo(ribbonelement, "gallery", "getLabel", nameof(CustomUI.GetLabel), RibbonBindingType.LabelBinding, null, extRibbondata, localizationPrefix));
                }
                return new Tuple<string, List<BindingInfo>>(root.ToString(), bindingInfos);
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return null;
        }
        #endregion

        #region private Functions
        private static List<BindingInfo> ReplaceBindingAndExtractBindingInfo(XElement ribbonElement, string controlType, string attributeName, string newAttributeValue, RibbonBindingType bindingType, string parentID = null, IExtendedRibbonData extRibbondata = null, string localizationPrefix = "")
        {
            List<BindingInfo> bindingInfos = new List<BindingInfo>();
            var elements = ribbonElement.DescendantsAndSelf().Where(ele => ele.Name.LocalName == controlType || controlType == "").ToList();
            foreach (var element in elements)
            {
                var attr = element.Attributes().FirstOrDefault(atr => atr.Name.LocalName == attributeName);
                if (attr != null)
                {
                    BindingInfo newBindingInfo = null;
                    if (attr.Value.StartsWith("{Binding"))
                    {
                        newBindingInfo = new BindingInfo()
                        {
                            BindingPath = attr.Value.Replace("{Binding", "").Replace("}", "").Trim(),
                            RibbonBindingType = bindingType,
                            ID = element.Attributes().FirstOrDefault(atr => atr.Name.LocalName == "id")?.Value ?? "",
                            ParentID = parentID
                        };
                        bindingInfos.Add(newBindingInfo);

                        if (controlType == "gallery" && newBindingInfo != null && attributeName == "itemssource")
                        {
                            newBindingInfo.SubInfos.AddRange(ReplaceBindingAndExtractBindingInfo(element, "gallery", "getItemID", nameof(CustomUI.GetItemID), RibbonBindingType.ItemId, newBindingInfo.ID));
                            newBindingInfo.SubInfos.AddRange(ReplaceBindingAndExtractBindingInfo(element, "gallery", "getItemLabel", nameof(CustomUI.GetItemLabel), RibbonBindingType.ItemLabel, newBindingInfo.ID));
                            newBindingInfo.SubInfos.AddRange(ReplaceBindingAndExtractBindingInfo(element, "gallery", "getItemImage", nameof(CustomUI.GetItemImage), RibbonBindingType.ItemImage, newBindingInfo.ID));
                        }
                    }
                    else
                    {
                        if (attr.Value.StartsWith("{Loc "))
                        {
                            var resourceKey = attr.Value.Replace("{Loc ", "").Replace("}", "").Trim();

                            newBindingInfo = new BindingInfo()
                            {
                                ResourceKey = localizationPrefix + resourceKey,
                                RibbonBindingType = RibbonBindingType.LabelFromResource,
                                ID = element.Attributes().FirstOrDefault(atr => atr.Name.LocalName == "id")?.Value ?? "",
                                ParentID = parentID
                            };
                            bindingInfos.Add(newBindingInfo);
                        }
                    }
                    if (newAttributeValue == "#remove")
                    {
                        attr.Remove();
                    }
                    else
                    {
                        attr.Value = newAttributeValue;
                    }
                }
            }
            return bindingInfos;
        }
        #endregion

    }
}
