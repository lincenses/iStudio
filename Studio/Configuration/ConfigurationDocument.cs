using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio.Configuration
{
    public class ConfigurationDocument
    {
        #region 公有成员
        public const string XmlRootName = "configuration";              // Xml文件根节点名称。
        public const string XmlGroupNodeName = "group";                 // Xml文件组节点名称。
        public const string XmlItemNodeName = "item";                   // Xml文件项节点名称。
        public const string XmlKeyAttributeName = "key";                // Xml文件Key属性的名称。
        public const string XmlValueAttributeName = "value";            // Xml文件value属性的名称。
        public const string XmlEncryptedAttributeName = "encrypted";    // Xml文件Encrypted属性的名称。
        #endregion

        #region 私有成员
        private IDictionary<string, ConfigurationGroup> _ConfigurationGroupDictionary;
        #endregion

        #region 公有属性

        #region 获取具有指定键的配置组
        public ConfigurationGroup this[string key] => _ConfigurationGroupDictionary.ContainsKey(key) ? _ConfigurationGroupDictionary[key] : null;
        #endregion

        #region 获取所包含的配置组数量
        public int Count => _ConfigurationGroupDictionary.Count;
        #endregion

        #endregion

        #region 构造函数
        public ConfigurationDocument()
        {
            _ConfigurationGroupDictionary = new Dictionary<string, ConfigurationGroup>(StringComparer.OrdinalIgnoreCase);
        }
        #endregion

        #region 公有方法

        #region 确定是否包换具有指定键的配置组
        public bool Contains(string key)
        {
            return _ConfigurationGroupDictionary.ContainsKey(key);
        }
        #endregion

        #region 返回一个循环访问集合的枚举器
        public IEnumerator<ConfigurationGroup> GetEnumerator()
        {
            return _ConfigurationGroupDictionary.Values.GetEnumerator();
        }
        #endregion

        #region 移除具有指定键的配置组
        public bool Remove(string key)
        {
            return _ConfigurationGroupDictionary.Remove(key);
        }
        #endregion

        #region 清空所有配置组
        public void Clear()
        {
            _ConfigurationGroupDictionary.Clear();
        }
        #endregion

        #region 添加并返回具有指定键的配置组
        public ConfigurationGroup Add(string key)
        {
            ConfigurationGroup value = new ConfigurationGroup(key, this);
            _ConfigurationGroupDictionary.Add(key, value);
            return value;
        }
        #endregion

        #region 获取具有指定键的配置组，如果不存在则在配置文件中创建具有指定键的配置组并返回该配置组
        public ConfigurationGroup GetOrAdd(string key)
        {
            return _ConfigurationGroupDictionary.ContainsKey(key) ? _ConfigurationGroupDictionary[key] : Add(key);
        }
        #endregion

        #region 获取具有指定键的配置项的值，如果不存在则返回空字符串
        public string GetConfigurationItemValue(string groupKey, string itemKey)
        {
            if (_ConfigurationGroupDictionary.ContainsKey(groupKey))
            {
                return _ConfigurationGroupDictionary[groupKey].GetConfigurationItemValue(itemKey);
            }
            return "";
        }
        #endregion

        #region 保存到文件
        public void Save(string fileName)
        {
            // 创建文件对象。
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(fileName);
            // 如果文件对象所在目录不存在，则创建。
            if (!System.IO.Directory.Exists(fileInfo.DirectoryName))
            {
                System.IO.Directory.CreateDirectory(fileInfo.DirectoryName);
            }
            // 创建Xml文档。
            System.Xml.Linq.XDocument xDocument = new System.Xml.Linq.XDocument(new System.Xml.Linq.XDeclaration("1.0", "utf-8", ""));
            // 添加Xml根节点。
            xDocument.Add(new System.Xml.Linq.XElement(XmlRootName));
            // 遍历配置文件中的配置组集合。
            foreach (ConfigurationGroup configurationGroup in _ConfigurationGroupDictionary.Values)
            {
                // 创建并添加Xml组节点。
                System.Xml.Linq.XElement xGroup = new System.Xml.Linq.XElement(XmlGroupNodeName);
                xGroup.SetAttributeValue(XmlKeyAttributeName, configurationGroup.Key);
                xDocument.Root.Add(xGroup);
                // 遍历配置组中的配置项集合。
                foreach (ConfigurationItem configurationItem in configurationGroup)
                {
                    // 创建并添加Xml项节点。
                    System.Xml.Linq.XElement xItem = new System.Xml.Linq.XElement(XmlItemNodeName);
                    xItem.SetAttributeValue(XmlKeyAttributeName, configurationItem.Key);
                    xItem.SetAttributeValue(XmlValueAttributeName, configurationItem.Encrypted ? Studio.Security.DESManager.Encrypt(configurationItem.Value) : configurationItem.Value);
                    xItem.SetAttributeValue(XmlEncryptedAttributeName, configurationItem.Encrypted);
                    xGroup.Add(xItem);
                }
            }
            // 保存到文件。
            xDocument.Save(fileInfo.FullName);
        }
        #endregion

        #region 从文件中加载配置文件
        /// <summary>
        /// 加载配置文件。
        /// </summary>
        /// <param name="fileName">文件名。</param>
        /// <returns>包含所指定文件内容的配置文件。</returns>
        public static ConfigurationDocument LoadFrom(string fileName)
        {
            // 创建文件对象。
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(fileName);
            // 如果文件不存在，则返回空的配置文件对象。
            if (!fileInfo.Exists)
            {
                return new ConfigurationDocument();
            }
            // 从文件中加载Xml文档对象。
            System.Xml.Linq.XDocument xDocument = System.Xml.Linq.XDocument.Load(fileInfo.FullName);
            // 如果Xml文档的根节点名称不是默认的根节点名称，则返回空的配置文件对象。
            if (xDocument.Root.Name.LocalName != XmlRootName)
            {
                return new ConfigurationDocument();
            }
            // 创建配置文件对象。
            ConfigurationDocument document = new ConfigurationDocument();
            // 遍历Xml文档根节点下的Xml组节点。
            foreach (System.Xml.Linq.XElement xGroup in xDocument.Root.Elements())
            {
                // 如果Xml组节点的名称是默认的组节点的名称，则继续。
                if (xGroup.Name.LocalName == XmlGroupNodeName)
                {
                    // 获取Xml组节点的key属性。
                    System.Xml.Linq.XAttribute xGroupKeyAttribute = xGroup.Attribute(XmlKeyAttributeName);
                    // 如果key属性不为空且属性值不为空，则继续。
                    if (xGroupKeyAttribute != null)
                    {
                        // 根据xml组节点的key属性值创建并添加配置组对象。
                        ConfigurationGroup configurationGroup = document.GetOrAdd(xGroupKeyAttribute.Value);
                        // 遍历Xml组节点下的Xml项节点。
                        foreach (System.Xml.Linq.XElement xItem in xGroup.Elements())
                        {
                            // 如果Xml项节点的名称是默认的项节点名称，则继续。
                            if (xItem.Name.LocalName == XmlItemNodeName)
                            {
                                // 获取Xml项节点的key属性、value属性、encrypted属性。
                                System.Xml.Linq.XAttribute xItemKeyAttribute = xItem.Attribute(XmlKeyAttributeName);
                                System.Xml.Linq.XAttribute xItemValueAttribute = xItem.Attribute(XmlValueAttributeName);
                                System.Xml.Linq.XAttribute xItemEncryptedAttribute = xItem.Attribute(XmlEncryptedAttributeName);
                                // 如果key属性不为空，且key属性值不为空
                                if (xItemKeyAttribute != null)
                                {
                                    // 创建配置项。
                                    ConfigurationItem configurationItem = configurationGroup.GetOrAdd(xItemKeyAttribute.Value);
                                    // 如果value属性不为空，则赋值给配置项。
                                    if (xItemValueAttribute != null)
                                    {
                                        configurationItem.Value = xItemValueAttribute.Value;
                                    }
                                    // 如果encrypted属性不为空，且encrypted属性值不为空，则解密配置项的值。
                                    if (xItemEncryptedAttribute != null && !string.IsNullOrWhiteSpace(xItemEncryptedAttribute.Value))
                                    {
                                        // 如果encrypted属性可以转换成bool类型的值。
                                        bool encrypted = false;
                                        if (bool.TryParse(xItemEncryptedAttribute.Value, out encrypted))
                                        {
                                            // 赋值配置项的加密属性。
                                            configurationItem.Encrypted = encrypted;
                                            // 根据配置项的加密属性值来解密配置项的值。
                                            if (encrypted)
                                            {
                                                if (Studio.Security.DESManager.TryDecrypt(configurationItem.Value, out string value))
                                                {
                                                    configurationItem.Value = value;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return document;
        }
        #endregion

        #endregion
    }
}
