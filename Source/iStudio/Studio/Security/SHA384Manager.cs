using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio.Security
{
    public class SHA384Manager
    {
        #region SHA384加密
        public static string Encrypt(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) { return ""; }
            byte[] byteArray = Encoding.UTF8.GetBytes(text);
            using (System.Security.Cryptography.SHA384CryptoServiceProvider shaCSP = new System.Security.Cryptography.SHA384CryptoServiceProvider())
            {
                return BitConverter.ToString(shaCSP.ComputeHash(byteArray)).Replace("-", "");
            }
        }
        #endregion
    }
}
