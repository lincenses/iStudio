using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio.Security
{
    public class SHA1Manager
    {
        #region SHA1加密
        public static string Encrypt(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) { return ""; }
            byte[] byteArray = Encoding.UTF8.GetBytes(text);
            using (System.Security.Cryptography.SHA1CryptoServiceProvider shaCSP = new System.Security.Cryptography.SHA1CryptoServiceProvider())
            {
                return BitConverter.ToString(shaCSP.ComputeHash(byteArray)).Replace("-", "");
            }
        }
        #endregion
    }
}
