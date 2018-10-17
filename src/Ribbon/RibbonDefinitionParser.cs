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
        public static Tuple<string, List<BindingInfo>> ParseDefinition(string ribbonDefinitionText, IAddInInformation addInInfo)
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
                    bindingInfos.AddRange(ReplaceBindingAndExtractBindingInfo(ribbonelement, "button", "getLabel", nameof(CustomUI.GetLabel), RibbonBindingType.LabelBinding, null, addInInfo, localizationPrefix));
                    bindingInfos.AddRange(ReplaceBindingAndExtractBindingInfo(ribbonelement, "toggleButton", "getLabel", nameof(CustomUI.GetLabel), RibbonBindingType.LabelBinding, null, addInInfo, localizationPrefix));
                    bindingInfos.AddRange(ReplaceBindingAndExtractBindingInfo(ribbonelement, "gallery", "getLabel", nameof(CustomUI.GetLabel), RibbonBindingType.LabelBinding, null, addInInfo, localizationPrefix));
                    bindingInfos.AddRange(ReplaceBindingAndExtractBindingInfo(ribbonelement, "button", "getImage", nameof(CustomUI.GetImage), RibbonBindingType.ImageBinding, null, addInInfo, localizationPrefix));
                    bindingInfos.AddRange(ReplaceBindingAndExtractBindingInfo(ribbonelement, "", "getScreentip", nameof(CustomUI.GetScreentip), RibbonBindingType.ScreentipBinding, null, addInInfo, localizationPrefix));
                    bindingInfos.AddRange(ReplaceBindingAndExtractBindingInfo(ribbonelement, "", "getSupertip", nameof(CustomUI.GetSupertip), RibbonBindingType.SupertipBinding, null, addInInfo, localizationPrefix));
                    bindingInfos.AddRange(ReplaceBindingAndExtractBindingInfo(ribbonelement, "toggleButton", "getImage", nameof(CustomUI.GetImage), RibbonBindingType.ImageBinding, null, addInInfo, localizationPrefix));
                    bindingInfos.AddRange(ReplaceBindingAndExtractBindingInfo(ribbonelement, "gallery", "getImage", nameof(CustomUI.GetImage), RibbonBindingType.ImageBinding, null, addInInfo, localizationPrefix));
                    bindingInfos.AddRange(ReplaceBindingAndExtractBindingInfo(ribbonelement, "group", "getLabel", nameof(CustomUI.GetLabel), RibbonBindingType.LabelBinding, null, addInInfo, localizationPrefix));
                    bindingInfos.AddRange(ReplaceBindingAndExtractBindingInfo(ribbonelement, "comboBox", "onChange", nameof(CustomUI.OnSelectedChanged), RibbonBindingType.ComboboxSelectedChanged, null, addInInfo, localizationPrefix));
                    bindingInfos.AddRange(ReplaceBindingAndExtractBindingInfo(ribbonelement, "comboBox", "getText", nameof(CustomUI.GetText), RibbonBindingType.ComboboxSelectedText, null, addInInfo, localizationPrefix));
                    bindingInfos.AddRange(ReplaceBindingAndExtractBindingInfo(ribbonelement, "group", "getLabel", nameof(CustomUI.GetLabel), RibbonBindingType.LabelBinding, null, addInInfo, localizationPrefix));
                    bindingInfos.AddRange(ReplaceBindingAndExtractBindingInfo(ribbonelement, "comboBox", "itemssource", "#remove", RibbonBindingType.GalleryItemsSource));
                    bindingInfos.AddRange(ReplaceBindingAndExtractBindingInfo(ribbonelement, "labelControl", "getLabel", nameof(CustomUI.GetLabel), RibbonBindingType.LabelBinding, null, addInInfo, localizationPrefix));
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
        private static List<BindingInfo> ReplaceBindingAndExtractBindingInfo(XElement ribbonElement, string controlType, string attributeName, string newAttributeValue, RibbonBindingType bindingType, string parentID = null, IAddInInformation extRibbondata = null, string localizationPrefix = "")
        {
            List<BindingInfo> bindingInfos = new List<BindingInfo>();
            var elements = ribbonElement.DescendantsAndSelf().Where(ele => ele.Name.LocalName == controlType || controlType == "").ToList();
            foreach (var element in elements)
            {
                var attr = element.Attributes().FirstOrDefault(atr => atr.Name.LocalName == attributeName);
                if (attr != null)
                {
                    BindingInfo newBindingInfo = null;
                    string attrValueLower = attr.Value.ToLowerInvariant();
                    if (attrValueLower.StartsWith("{binding"))
                    {
                        newBindingInfo = new BindingInfo()
                        {
                            BindingPath = attr.Value.Substring(9).Replace("}", "").Trim(),
                            RibbonBindingType = bindingType,
                            ID = element.Attributes().FirstOrDefault(atr => atr.Name.LocalName == "id")?.Value ?? "",
                            ParentID = parentID
                        };
                        bindingInfos.Add(newBindingInfo);

                        if ((controlType == "gallery" || controlType == "comboBox") && newBindingInfo != null && attributeName == "itemssource")
                        {
                            newBindingInfo.SubInfos.AddRange(ReplaceBindingAndExtractBindingInfo(element, "gallery", "getItemID", nameof(CustomUI.GetItemID), RibbonBindingType.ItemId, newBindingInfo.ID));
                            newBindingInfo.SubInfos.AddRange(ReplaceBindingAndExtractBindingInfo(element, "gallery", "getItemLabel", nameof(CustomUI.GetItemLabel), RibbonBindingType.ItemLabel, newBindingInfo.ID));
                            newBindingInfo.SubInfos.AddRange(ReplaceBindingAndExtractBindingInfo(element, "gallery", "getItemImage", nameof(CustomUI.GetItemImage), RibbonBindingType.ItemImage, newBindingInfo.ID));

                            newBindingInfo.SubInfos.AddRange(ReplaceBindingAndExtractBindingInfo(element, "comboBox", "getItemID", nameof(CustomUI.GetItemID), RibbonBindingType.ItemId, newBindingInfo.ID));
                            newBindingInfo.SubInfos.AddRange(ReplaceBindingAndExtractBindingInfo(element, "comboBox", "getItemLabel", nameof(CustomUI.GetItemLabel), RibbonBindingType.ItemLabel, newBindingInfo.ID));
                            newBindingInfo.SubInfos.AddRange(ReplaceBindingAndExtractBindingInfo(element, "comboBox", "getItemImage", nameof(CustomUI.GetItemImage), RibbonBindingType.ItemImage, newBindingInfo.ID));

                            if (!element.Attributes().Any(atr => atr.Name.LocalName == "getItemCount"))
                            {
                                element.Add(new XAttribute("getItemCount", nameof(CustomUI.GetItemCount)));
                            }
                        }
                    }
                    else
                    {
                        if (attrValueLower.StartsWith("{res "))
                        {
                            var resourceKey = attr.Value.Substring(5).Replace("}", "").Trim();

                            newBindingInfo = new BindingInfo()
                            {
                                ResourceKey = localizationPrefix + resourceKey,
                                ID = element.Attributes().FirstOrDefault(atr => atr.Name.LocalName == "id")?.Value ?? "",
                                ParentID = parentID
                            };
                            switch (bindingType)
                            {
                                case RibbonBindingType.LabelBinding:
                                    newBindingInfo.RibbonBindingType = RibbonBindingType.LabelFromResource;
                                    break;
                                case RibbonBindingType.ImageBinding:
                                    newBindingInfo.RibbonBindingType = RibbonBindingType.ImageFromResource;
                                    break;
                                case RibbonBindingType.ScreentipBinding:
                                    newBindingInfo.RibbonBindingType = RibbonBindingType.ScreentipFromResource;
                                    break;
                                case RibbonBindingType.SupertipBinding:
                                    newBindingInfo.RibbonBindingType = RibbonBindingType.SupertipFromResource;
                                    break;
                                default:
                                    break;
                            }
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
