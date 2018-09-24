namespace ExcelDna_MVVM.Document
{

    #region Usings
    using Newtonsoft.Json.Linq;
    using NLog;
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Xml.Linq;
    #endregion

    public class SeDocument
    {
        private const string TABLE_ELEMENT_NAME = "seTables";
        #region LoggerInit
        private static Logger logger = LogManager.GetCurrentClassLogger();
        #endregion


        #region Properties & Variables
        object tablesLock = new object();
        object documentPropertiesLock = new object();
        private readonly Dictionary<string, (bool IsCustomDocumentProperty, string json, Dictionary<string, string> seData)> tables = new Dictionary<string, (bool IsCustomDocumentProperty, string json, Dictionary<string, string> seData)>();
        private readonly Dictionary<string, string> documentProperties = new Dictionary<string, string>();
        public dynamic Workbook { get; set; }
        private object CustomXMLPartsLock = new object();
        public string DocumentID
        {
            get { return Workbook.Name; }
        }
        #endregion

        #region Public Functions
        public string GetTable(string guid)
        {
            string data = null;
            if (tables.ContainsKey(guid))
            {
                var value = tables[guid];
                data = value.json;
            }
            return data;
        }

        public void SetTableData(string tableGuid, string key, string value)
        {
            try
            {
                if (key == "id" || key == "data")
                    throw new Exception("Keys 'data' and 'id' are not allowed");
                bool mustsave = false;
                lock (tablesLock)
                {
                    if (tables.ContainsKey(tableGuid))
                    {
                        var entry = tables[tableGuid];
                        if (!entry.seData.ContainsKey(key))
                        {
                            entry.seData.Add(key, value);
                        }
                        entry.seData[key] = value;
                        tables[tableGuid] = entry;

                        mustsave = true;
                    }
                }
                if (mustsave)
                    SaveDocumentProperties();
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }

        public string GetTableData(string tableGuid, string key)
        {
            string retval = "";

            try
            {
                lock (tablesLock)
                {
                    if (tables.ContainsKey(tableGuid))
                    {
                        var entry = tables[tableGuid];
                        if (entry.seData.ContainsKey(key))
                        {
                            retval = entry.seData[key];
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return retval;
        }

        public void LoadTableJsons()
        {
            try
            {

                logger.Trace("Read CustomXMLParts");

                string tablesJson = "";
                var tablesPart = GetTablePart(Workbook);
                if (tablesPart != null)
                {
                    logger.Trace($"Parsing XMLPart: XML={tablesPart.XML}");
                    XElement ele = XElement.Parse(tablesPart.XML);
                    var hypercubedefNodes = ele.Descendants(TABLE_ELEMENT_NAME).ToList();
                    if (hypercubedefNodes.Count > 0)
                    {
                        tablesJson = hypercubedefNodes[0].Value;
                    }

                    if (tablesJson != "")
                    {
                        dynamic dyntables = JObject.Parse(tablesJson);
                        foreach (var property in dyntables.Properties())
                        {
                            if (property.Name != "tables")
                            {
                                lock (documentPropertiesLock)
                                {
                                    if (!documentProperties.ContainsKey(property.Name))
                                        documentProperties.Add(property.Name, property.Value.ToString());
                                }
                            }
                        }
                        lock (tablesLock)
                        {
                            foreach (var item in dyntables.tables)
                            {
                                string tableId = item?.id?.ToString() ?? "";
                                string tabledata = item?.data?.ToString() ?? "";
                                var seData = new Dictionary<string, string>();

                                foreach (var property in item.Properties())
                                {
                                    if (property.Name != "id" && property.Name != "data")
                                    {
                                        if (!seData.ContainsKey(property.Name))
                                            seData.Add(property.Name, property.Value.ToString());
                                    }
                                }

                                tables.Add(tableId, (false, tabledata, seData));
                            }
                        }
                    }
                }


                logger.Trace("Read CustomDocumentProperties");
                var props = Workbook.CustomDocumentProperties;
                lock (tablesLock)
                {
                    foreach (var worksheet in Workbook.Worksheets)
                    {
                        foreach (var tb in worksheet.ListObjects)
                        {
                            var guID = tb.Comment;
                            string json = ReadLongString(props, guID).ToString();
                            logger.Trace($"found Value in documentpropertyies for Key:{guID} json:{json}");
                            if (!tables.ContainsKey(guID))
                                tables.Add(guID, (true, json, new Dictionary<string, string>()));
                        }
                    }
                }
                if (SaveDocumentProperties())
                {
                    RemoveCustomDocumentProperties();
                }

            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }

        }

        public bool SaveOrUpdateTableJson(string guid, string tableJson)
        {
            try
            {
                lock (tablesLock)
                {
                    if (tables.ContainsKey(guid))
                    {
                        var entry = tables[guid];
                        entry.json = tableJson;
                        tables[guid] = entry;
                    }
                    else
                    {
                        tables.Add(guid, (false, tableJson, new Dictionary<string, string>()));
                    }
                }
                SaveDocumentProperties();
                return true;
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return false;

        }

        public bool SaveDocumentProperties()
        {
            try
            {
                dynamic dynTablesJson = new JObject();
                foreach (var prop in documentProperties)
                {
                    if (prop.Value != null)
                        dynTablesJson[prop.Key] = prop.Value;
                }

                lock (tablesLock)
                {
                    logger.Trace("setting tableData");
                    dynTablesJson.tables = new JArray();
                    foreach (var item in tables)
                    {
                        try
                        {
                            dynamic dynTable = new JObject();
                            dynTable.id = item.Key;
                            dynTable.data = JObject.Parse(item.Value.json);
                            string online = GetTableData(item.Key, "online");
                            foreach (var prop in item.Value.seData)
                            {
                                dynTable[prop.Key] = prop.Value;
                            }
                            dynTablesJson.tables.Add(dynTable);
                        }
                        catch (Exception ex)
                        {
                            logger.Error(ex);
                            logger.Trace($"Error {item.Key} - JSON - {item.Value.json} ");
                        }
                    }
                }
                lock (CustomXMLPartsLock)
                {

                    logger.Trace("Deleting tablespart");
                    int counter = 0;
                    try
                    {
                        var part = GetTablePart(Workbook);
                        while (part != null)
                        {
                            bool deleted = false;
                            try
                            {
                                logger.Trace($"try Delete {part.Id}");
                                part.Delete();
                                logger.Trace($"Deleted {part.Id}!");
                                deleted = true;
                            }
                            catch (Exception ex)
                            {
                                logger.Trace(ex);
                            }
                            if (!deleted)
                            {
                                Thread.Sleep(500);
                                counter++;
                                if (counter > 10)
                                {
                                    throw new Exception("More than 10 Trys failed! Adding second XMLPart.");
                                }
                            }
                            else
                            {
                                part = GetTablePart(Workbook);
                            }
                        }
                        //Thread.Sleep(500);
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex);
                    }

                    logger.Trace("Adding tablespart");
                    Workbook.CustomXMLParts.Add($"<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n<json>\r\n<{TABLE_ELEMENT_NAME}>\r\n<![CDATA[{dynTablesJson.ToString()}]]>\r\n</{TABLE_ELEMENT_NAME}></json>");
                }
                return true;
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return false;
        }

        public bool SetDocumentProperty(string propertyName, string value)
        {
            try
            {
                if (propertyName == "tables")
                    throw new Exception("Propertyname 'tables' is not allowed");
                lock (documentPropertiesLock)
                {
                    if (documentProperties.ContainsKey(propertyName))
                    {
                        documentProperties[propertyName] = value;
                    }
                    else
                    {
                        documentProperties.Add(propertyName, value);
                    }
                }
                return SaveDocumentProperties();
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return false;
        }

        public string GetDocumentProperty(string propertyName)
        {
            try
            {
                lock (documentPropertiesLock)
                {
                    if (documentProperties.ContainsKey(propertyName))
                    {
                        return documentProperties[propertyName];
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return null;
        }
        #endregion

        #region private Functions       
        private int String2Int(string s)
        {
            if (String.IsNullOrEmpty(s))
                return 0;

            try
            {
                return int.Parse(s.Substring(1));
            }
            catch
            {
                return 0;
            }
        }

        private string ReadLongString(dynamic dp, string key)
        {
            if (string.IsNullOrWhiteSpace(key))
                return null;

            //var keyList = (from c in dp where c.Name.StartsWith(key) || c.Name.StartsWith(key + "_") orderby String2Int(c.Name.Substring(key.Length)) select c.Value).ToList();
            Dictionary<string, string> documentproperties = new Dictionary<string, string>();
            foreach (var c in dp)
            {
                documentproperties.Add(c.Name, c.Value.ToString());
                logger.Trace($"found documentproperty for Key:{key} Name:{c.Name} Value:{c.Value.ToString()}");
            }
            var keyList = documentproperties.Where(kvPair => kvPair.Key.StartsWith(key) || kvPair.Key.StartsWith(key + "_"))
                .OrderBy(kvPair => String2Int(kvPair.Key.Substring(kvPair.Key.Length)))
                .Select(kvPair => kvPair.Value)
                .ToList();

            if (keyList.Count > 0)
            {
                string res = "";
                foreach (var item in keyList)
                    res += item;

                return res;
            }

            return null;
        }

        private void RemoveCustomDocumentProperties()
        {
            try
            {
                lock (CustomXMLPartsLock)
                {
                    logger.Trace("Reading tablespart");
                    dynamic dp = Workbook.CustomDocumentProperties;
                    if (dp != null)
                    {
                        lock (tablesLock)
                        {
                            foreach (var item in tables)
                            {
                                foreach (var prop in dp)
                                {
                                    if (prop.Name.StartsWith(item.Key))
                                    {
                                        logger.Trace($"Deleting DocumentProperties: {prop.Name}");
                                        prop.Delete();
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

        private dynamic GetTablePart(dynamic workbook)
        {
            //var tablepart = workbook.CustomXMLParts.FirstOrDefault(part => part.XML.Contains(TABLE_ELEMENT_NAME));
            dynamic tablepart = null;
            try
            {
                lock (CustomXMLPartsLock)
                {
                    logger.Trace("Reading tablespart 2");
                    foreach (var part in workbook.CustomXMLParts)
                    {
                        if (part.XML.Contains(TABLE_ELEMENT_NAME))
                        {
                            tablepart = part;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
            return tablepart;
        }
        #endregion



    }
}
