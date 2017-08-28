using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio.Security
{
    public class Admin
    {
        #region 公有成员
        // 密码件名。
        public readonly static string FileName = System.Windows.Forms.Application.StartupPath + @"\Admin.dat";
        #endregion

        #region 私有方法

        #region 从文件中获取密码
        private static string LoadPasswordFromFile(string fileName)
        {
            string value = "";
            List<byte> byteList = new List<byte>();
            System.IO.FileStream stream = new System.IO.FileStream(FileName, System.IO.FileMode.Open, System.IO.FileAccess.Read);
            System.IO.BinaryReader reader = new System.IO.BinaryReader(stream);
            while (reader.PeekChar() != -1)
            {
                byteList.Add(reader.ReadByte());
            }
            reader.Close();
            reader.Dispose();
            stream.Close();
            stream.Dispose();
            value = Encoding.UTF8.GetString(byteList.ToArray());
            value = value.Replace(Environment.NewLine, "");
            value = value.Replace(" ", "");
            value = DESManager.Decrypt(value);
            value = DESManager.Decrypt(value);
            value = DESManager.Decrypt(value);
            return value;
        }
        #endregion

        #region 将密码保存到文件
        private static void SavePasswordToFile(string password, string fileName)
        {
            string value = DESManager.Encrypt(password);
            value = DESManager.Encrypt(value);
            value = DESManager.Encrypt(value);
            value = string.Join(" ", System.Text.RegularExpressions.Regex.Matches(value, @"..").Cast<System.Text.RegularExpressions.Match>().ToList());
            List<string> valueList = new List<string>();
            int i = 0;
            int count = value.Length / 36;
            for (i = 0; i < count; i++)
            {
                valueList.Add(value.Substring(i * 36, 36));
            }
            valueList.Add(value.Substring(i * 36));
            value = string.Join(Environment.NewLine, valueList);
            byte[] buff = Encoding.UTF8.GetBytes(value);
            System.IO.FileStream stream = stream = new System.IO.FileStream(fileName, System.IO.FileMode.Create, System.IO.FileAccess.ReadWrite);
            System.IO.BinaryWriter writer = new System.IO.BinaryWriter(stream);
            writer.Write(buff);
            writer.Close();
            writer.Dispose();
            stream.Close();
            stream.Dispose();
        }
        #endregion

        #endregion

        #region 公有方法

        #region 验证密码是否和密码文件中的密码一致
        public static bool VerifyPassword(string password)
        {
            return GetPassword() == password;
        }
        #endregion

        #region 获取默认密码
        public static string GetDefaulePassword()
        {
            string value = DateTime.Now.ToString("yyyyMMdd");
            value = MD5Manager.Encrypt16(value).Replace("-", "");
            return value.Substring(4, 8);
        }
        #endregion

        #region 获取默认密码文件中的密码，如果密码文件不存在，会自动创建具有默认密码的密码文件
        public static string GetPassword()
        {
            if (!System.IO.File.Exists(FileName))
            { SavePassword(GetDefaulePassword()); }
            return LoadPasswordFromFile(FileName);
        }
        #endregion

        #region 从指定的密码文件中获取密码
        public static string GetPassword(string fileName)
        {
            return LoadPasswordFromFile(fileName);
        }
        #endregion

        #region 将密码保存到默认密码文件
        public static void SavePassword(string password)
        {
            SavePasswordToFile(password, FileName);
        }
        #endregion

        #region 将密码保存到指定的密码文件
        public static void SavePassword(string password, string fileName)
        {
            SavePasswordToFile(password, fileName);
        }
        #endregion

        #endregion
    }
}
