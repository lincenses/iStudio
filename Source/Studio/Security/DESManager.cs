using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio.Security
{
    public class DESManager
    {
        #region 私有成员
        private static readonly byte[] KEY = { 0x01, 0x23, 0x45, 0x67, 0x89, 0xAB, 0xCD, 0xEF };  // 默认密钥
        private static readonly byte[] IV = { 0xFE, 0xDC, 0xBA, 0x98, 0x76, 0x54, 0x32, 0x10 };  // 默认密钥向量
        #endregion

        #region DES加密
        public static string Encrypt(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) { return ""; }
            byte[] byteArray = Encoding.UTF8.GetBytes(text);
            using (System.Security.Cryptography.DESCryptoServiceProvider desCSP = new System.Security.Cryptography.DESCryptoServiceProvider())
            {
                using (System.IO.MemoryStream memoryStream = new System.IO.MemoryStream())
                {
                    System.Security.Cryptography.ICryptoTransform iCryptoTransform = desCSP.CreateEncryptor(KEY, IV);
                    using (System.Security.Cryptography.CryptoStream cryptoStream = new System.Security.Cryptography.CryptoStream(memoryStream, iCryptoTransform, System.Security.Cryptography.CryptoStreamMode.Write))
                    {
                        cryptoStream.Write(byteArray, 0, byteArray.Length);
                        cryptoStream.FlushFinalBlock();
                    }
                    return BitConverter.ToString(memoryStream.ToArray()).Replace("-", "");
                }
            }
        }
        #endregion

        #region DES解密
        public static string Decrypt(string text)
        {
            string inputText = text;
            if (string.IsNullOrWhiteSpace(text))
            {
                return "";
            }
            if (inputText.Contains("-"))
            {
                inputText = inputText.Replace("-", "");
            }
            if (inputText.Contains(" "))
            {
                inputText = inputText.Replace(" ", "");
            }
            if (inputText.Length % 2 != 0)
            {
                throw new ArgumentException("要解密的字符串长度无效。");
            }
            byte[] byteArray = new byte[inputText.Length / 2];
            for (int i = 0; i < byteArray.Length; i++)
            {
                byteArray[i] = byte.Parse(inputText.Substring(i * 2, 2), System.Globalization.NumberStyles.HexNumber);
            }
            using (System.Security.Cryptography.DESCryptoServiceProvider desCSP = new System.Security.Cryptography.DESCryptoServiceProvider())
            {
                using (System.IO.MemoryStream memoryStream = new System.IO.MemoryStream())
                {
                    System.Security.Cryptography.ICryptoTransform iCryptoTransform = desCSP.CreateDecryptor(KEY, IV);
                    using (System.Security.Cryptography.CryptoStream cryptoStream = new System.Security.Cryptography.CryptoStream(memoryStream, iCryptoTransform, System.Security.Cryptography.CryptoStreamMode.Write))
                    {
                        cryptoStream.Write(byteArray, 0, byteArray.Length);
                        cryptoStream.FlushFinalBlock();
                    }
                    return Encoding.UTF8.GetString(memoryStream.ToArray());
                }
            }
        }

        public static bool TryDecrypt(string text, out string value)
        {
            try
            {
                value = Decrypt(text);
                return true;
            }
            catch
            {
                value = "";
                return false;
            }
        }
        #endregion

    }
}
