using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio.Configuration
{
    public class ConfigurationGroup
    {
        #region 私有成员
        private IDictionary<string, ConfigurationItem> _ConfigurationItemDictionary;
        #endregion

        #region 公有属性

        #region 获取配置组的键
        public string Key { get; private set; } = "";
        #endregion

        #region 获取所隶属的配置文件对象
        public ConfigurationDocument Parent { get; private set; } = null;
        #endregion

        #region 获取具有指定键的配置项
        public ConfigurationItem this[string key] => _ConfigurationItemDictionary.ContainsKey(key) ? _ConfigurationItemDictionary[key] : null;
        #endregion

        #region 获取配置项的数量
        public int Count => _ConfigurationItemDictionary.Count;
        #endregion

        #endregion

        #region 构造函数
        internal ConfigurationGroup(string key, ConfigurationDocument parent)
        {
            _ConfigurationItemDictionary = new Dictionary<string, ConfigurationItem>(StringComparer.OrdinalIgnoreCase);
            Key = key;
            Parent = parent;
        }
        #endregion

        #region 公有方法

        #region 确定是否包含具有指定键的配置项
        public bool Contains(string key)
        {
            return _ConfigurationItemDictionary.ContainsKey(key);
        }
        #endregion

        #region 返回一个循环访问集合的枚举
        public IEnumerator<ConfigurationItem> GetEnumerator()
        {
            return _ConfigurationItemDictionary.Values.GetEnumerator();
        }
        #endregion

        #region 移除具有指定键的配置项
        public bool Remove(string key)
        {
            return _ConfigurationItemDictionary.Remove(key);
        }
        #endregion

        #region 清空所有配置项
        public void Clear()
        {
            _ConfigurationItemDictionary.Clear();
        }
        #endregion

        #region 添加并返回具有指定键的配置项
        public ConfigurationItem Add(string key)
        {
            return Add(key, "", false);
        }
        #endregion

        #region 添加并返回具有指定键和值的配置项
        public ConfigurationItem Add(string key, string value)
        {
            return Add(key, value, false);
        }
        #endregion

        #region 添加并返回具有指定键、值和是否加密的配置项
        public ConfigurationItem Add(string key, string value, bool encrypted)
        {
            ConfigurationItem item = new ConfigurationItem(key, value, encrypted, this);
            _ConfigurationItemDictionary.Add(key, item);
            return item;
        }
        #endregion

        #region 获取具有指定键的配置项，如果不存在则在所隶属的配置组中添加具有指定键的配置项并返回该配置项
        public ConfigurationItem GetOrAdd(string key)
        {
            return _ConfigurationItemDictionary.ContainsKey(key) ? _ConfigurationItemDictionary[key] : Add(key);
        }
        #endregion

        #region 获取具有指定键的配置项的值，如果配置项不存在则返回空字符串
        public string GetConfigurationItemValue(string key)
        {
            return _ConfigurationItemDictionary.ContainsKey(key) ? _ConfigurationItemDictionary[key].Value : "";
        }
        #endregion

        #endregion
    }
}
