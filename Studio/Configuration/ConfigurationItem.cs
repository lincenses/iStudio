using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio.Configuration
{
    public class ConfigurationItem
    {
        #region 公有属性

        #region 获取配置项的键
        public string Key { get; private set; } = "";
        #endregion

        #region 获取配置项的值
        public string Value { get; set; } = "";
        #endregion

        #region 获取配置项是否加密
        public bool Encrypted { get; set; } = false;
        #endregion

        #region 获取所隶属的配置组对象
        public ConfigurationGroup Parent { get; private set; } = null;
        #endregion

        #endregion

        #region 构造函数
        internal ConfigurationItem(string key, string value, bool encrypted, ConfigurationGroup parent)
        {
            Key = key;
            Value = value;
            Encrypted = encrypted;
            Parent = parent;
        }
        #endregion
    }
}
