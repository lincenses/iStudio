using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio.Configuration
{
    public class DatabaseConfiguration
    {
        #region 公有属性

        #region 获取或设置服务器
        public string DataSource { get; set; } = "";
        #endregion

        #region 获取或设置数据库
        public string InitialCatalog { get; set; } = "";
        #endregion

        #region 获取或设置登陆名
        public string UserID { get; set; } = "";
        #endregion

        #region 获取或设置密码
        public string Password { get; set; } = "";
        #endregion

        #region 获取数据库连接字符串
        public string ConnectionString => GetConnectionStringBuilder().ConnectionString;
        #endregion

        #endregion

        #region 构造函数
        public DatabaseConfiguration() : this("", "", "", "") { }

        public DatabaseConfiguration(string dataSource, string initialCatalog, string userID, string password)
        {
            DataSource = dataSource;
            InitialCatalog = initialCatalog;
            UserID = userID;
            Password = password;
        }
        #endregion

        #region 公有方法

        #region 加载数据库配置文件
        public void Load(string fileName)
        {
            string fullName = new System.IO.FileInfo(fileName).FullName;
            ConfigurationDocument document = ConfigurationDocument.LoadFrom(fullName);
            DataSource = document.GetValue(GroupKeys.DataBase, ItemKeys.DataSource);
            InitialCatalog = document.GetValue(GroupKeys.DataBase, ItemKeys.InitialCatalog);
            UserID = document.GetValue(GroupKeys.DataBase, ItemKeys.UserID);
            Password = document.GetValue(GroupKeys.DataBase, ItemKeys.Password);
        }
        #endregion

        #region 加载数据库配置文件
        public static DatabaseConfiguration LoadFrom(string fileName)
        {
            string fullName = new System.IO.FileInfo(fileName).FullName;
            ConfigurationDocument document = ConfigurationDocument.LoadFrom(fullName);
            DatabaseConfiguration configuration = new DatabaseConfiguration()
            {
                DataSource = document.GetValue(GroupKeys.DataBase, ItemKeys.DataSource),
                InitialCatalog = document.GetValue(GroupKeys.DataBase, ItemKeys.InitialCatalog),
                UserID = document.GetValue(GroupKeys.DataBase, ItemKeys.UserID),
                Password = document.GetValue(GroupKeys.DataBase, ItemKeys.Password)
            };
            return configuration;
        }
        #endregion

        #region 将数据库配置保存到文件
        public void Save(string fileName, bool encrypted)
        {
            string fullName = new System.IO.FileInfo(fileName).FullName;
            ConfigurationDocument document = new ConfigurationDocument();
            ConfigurationGroup group = document.Add(GroupKeys.DataBase);
            group.Add(ItemKeys.DataSource, DataSource, encrypted);
            group.Add(ItemKeys.InitialCatalog, InitialCatalog, encrypted);
            group.Add(ItemKeys.UserID, UserID, encrypted);
            group.Add(ItemKeys.Password, Password, encrypted);
            document.Save(fullName);
        }
        #endregion

        #region 从配置文件获取数据库连接字符串
        public System.Data.SqlClient.SqlConnectionStringBuilder GetConnectionStringBuilder()
        {
            System.Data.SqlClient.SqlConnectionStringBuilder sqlConnectionStringBuilder = new System.Data.SqlClient.SqlConnectionStringBuilder()
            {
                DataSource = DataSource,
                InitialCatalog = InitialCatalog,
                UserID = UserID,
                Password = Password
            };
            return sqlConnectionStringBuilder;
        }

        public System.Data.SqlClient.SqlConnectionStringBuilder GetConnectionStringBuilder(int connectTimeout)
        {
            System.Data.SqlClient.SqlConnectionStringBuilder sqlConnectionStringBuilder = new System.Data.SqlClient.SqlConnectionStringBuilder()
            {
                DataSource = DataSource,
                InitialCatalog = InitialCatalog,
                UserID = UserID,
                Password = Password,
                ConnectTimeout = connectTimeout
            };
            return sqlConnectionStringBuilder;
        }
        #endregion

        #endregion


        #region 配置组键的类
        public static class GroupKeys
        {
            public const string DataBase = @"Database";
        }
        #endregion

        #region 配置项键的类
        public static class ItemKeys
        {
            public const string DataSource = @"DataSource";
            public const string InitialCatalog = @"InitialCatalog";
            public const string UserID = @"UserID";
            public const string Password = @"Password";
        }
        #endregion

    }
}
