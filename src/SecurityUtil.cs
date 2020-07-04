//===================================================
//作    者：zhouzelong
//博    客：http://zhouzelong.cnblogs.com
//创建时间：2020-06-26 14:42:25
//备    注：
//===================================================

namespace USGame
{
    public static class SecurityUtil
    {
        private static readonly byte[] xorScale = new byte[] { 45, 66, 38, 55, 23, 254, 9, 165, 90, 19, 41, 45, 201, 58, 55, 37, 254, 185, 165, 169, 19, 171 };//.data文件的xor加解密因子

        /// <summary>
        /// 对数组进行异或
        /// </summary>
        /// <param name="buffer"></param>
        /// <returns></returns>
        public static byte[] Xor(byte[] buffer)
        {
            int iScaleLen = xorScale.Length;
            for (int i = 0; i < buffer.Length; i++)
            {
                buffer[i] = (byte)(buffer[i] ^ xorScale[i % iScaleLen]);
            }
            return buffer;
        }
    }
}