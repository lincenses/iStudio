using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio.Security
{
    public class SHA512Manager
    {
        #region SHA512加密
        public static string Encrypt(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) { return ""; }
            byte[] byteArray = Encoding.UTF8.GetBytes(text);
            using (System.Security.Cryptography.SHA512CryptoServiceProvider shaCSP = new System.Security.Cryptography.SHA512CryptoServiceProvider())
            {
                return BitConverter.ToString(shaCSP.ComputeHash(byteArray)).Replace("-", "");
            }
        }
        #endregion
    }
}
