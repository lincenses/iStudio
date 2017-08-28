using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio.Security
{
    public class MD5Manager
    {
        #region MD5加密（16位）
        public static string Encrypt16(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) { return ""; }
            byte[] byteArray = Encoding.UTF8.GetBytes(text);
            using (System.Security.Cryptography.MD5CryptoServiceProvider md5CSP = new System.Security.Cryptography.MD5CryptoServiceProvider())
            {
                return BitConverter.ToString(md5CSP.ComputeHash(byteArray), 4, 8).Replace("-", "");
            }
        }
        #endregion

        #region MD5加密（32位）
        public static string Encrypt32(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) { return ""; }
            byte[] byteArray = Encoding.UTF8.GetBytes(text);
            using (System.Security.Cryptography.MD5CryptoServiceProvider md5CSP = new System.Security.Cryptography.MD5CryptoServiceProvider())
            {
                return BitConverter.ToString(md5CSP.ComputeHash(byteArray)).Replace("-", "");
            }
        }
        #endregion
    }
}
