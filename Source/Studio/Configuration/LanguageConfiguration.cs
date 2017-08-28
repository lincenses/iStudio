using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio.Configuration
{
    public class LanguageConfiguration
    {
        #region 公有属性

        #region 获取或设置区域名称
        public string CultureName { get; set; } = "";
        #endregion

        #endregion

        #region 构造函数
        public LanguageConfiguration() : this("") { }

        public LanguageConfiguration(string cultureName)
        {
            CultureName = cultureName;
        }
        #endregion

        #region 公有方法

        #region 加载语言配置文件
        public void Load(string fileName)
        {
            string fullName = new System.IO.FileInfo(fileName).FullName;
            ConfigurationDocument document = ConfigurationDocument.LoadFrom(fullName);
            CultureName = document.GetValue(GroupKeys.Language, ItemKeys.CultureName);
        }
        #endregion

        #region 加载语言配置文件
        public static LanguageConfiguration LoadFrom(string fileName)
        {
            string fullName = new System.IO.FileInfo(fileName).FullName;
            ConfigurationDocument document = ConfigurationDocument.LoadFrom(fullName);
            LanguageConfiguration configuration = new LanguageConfiguration()
            {
                CultureName = document.GetValue(GroupKeys.Language, ItemKeys.CultureName)
            };
            return configuration;
        }
        #endregion

        #region 将语言配置保存到文件
        public void Save(string fileName, bool encrypted)
        {
            string fullName = new System.IO.FileInfo(fileName).FullName;
            ConfigurationDocument document = new ConfigurationDocument();
            document.Add(GroupKeys.Language).Add(ItemKeys.CultureName, CultureName, encrypted);
            document.Save(fullName);
        }
        #endregion

        #endregion



        #region 配置组键的类
        public static class GroupKeys
        {
            public const string Language = @"Language";
        }
        #endregion

        #region 配置项键的类
        public static class ItemKeys
        {
            public const string CultureName = @"CultureName";
        }
        #endregion
    }
}
