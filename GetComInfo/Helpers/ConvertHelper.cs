using System.Text;

namespace GetComInfo.Helpers
{
    /// <summary>
    /// Convert Helper class
    /// </summary>
    public static class ConvertHelper
    {
        /// <summary>
        /// Convert HexBytes to UnicodeString
        /// </summary>
        public static string HexBytes2UnicodeStr(byte[] ba)
        {
            var strMessage = Encoding.BigEndianUnicode.GetString(ba, 0, ba.Length);
            return strMessage;
        }

        /// <summary>
        /// Convert HexStr to HexBytes
        /// </summary>
        public static byte[] HexStr2HexBytes(string strHex)
        {
            strHex = strHex.Replace(" ", "");
            int nNumberChars = strHex.Length / 2;
            byte[] aBytes = new byte[nNumberChars];
            using (var sr = new StringReader(strHex))
            {
                for (int i = 0; i < nNumberChars; i++)
                {
                    try
                    {
                        aBytes[i] = Convert.ToByte(new String(new char[2] { (char)sr.Read(), (char)sr.Read() }), 16);
                    }
                    catch (Exception e)
                    {
                        return null;
                    }
                }
            }
            return aBytes;
        }
    }
}
