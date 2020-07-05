//===================================================
//作    者：zhouzelong
//博    客：http://zhouzelong.cnblogs.com
//创建时间：2020-06-26 14:27:37
//备    注：
//===================================================
using System;
using System.IO;

namespace USGame
{
    public static class IOUtil
    {
        #region GetFileText 读取文本文件
        /// <summary>
        /// 读取文本文件
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns></returns>
        public static string GetFileText(string filePath)
        {
            string content = string.Empty;

            if (!File.Exists(filePath))
            {
                return content;
            }

            using (StreamReader sr = File.OpenText(filePath))
            {
                content = sr.ReadToEnd();
            }
            return content;
        }
        #endregion

        #region CreateTextFile 创建文本文件
        /// <summary>
        /// 创建文本文件
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="content">文件内容</param>
        public static void CreateTextFile(string filePath, string content)
        {
            DeleteFile(filePath);

            using (FileStream fs = File.Create(filePath))
            {
                using (StreamWriter sw = new StreamWriter(fs))
                {
                    sw.Write(content.ToString());
                }
            }
        }
        #endregion

        #region DeleteFile 删除文件
        /// <summary>
        /// 删除文件
        /// </summary>
        /// <param name="filePath"></param>
        public static void DeleteFile(string filePath)
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }
        #endregion

        #region CopyDirectory 拷贝文件夹
        /// <summary>
        /// 拷贝文件夹
        /// </summary>
        /// <param name="sourceDirName">源文件夹</param>
        /// <param name="destDirName">目标文件夹</param>
        public static void CopyDirectory(string sourceDirName, string destDirName)
        {
            try
            {
                if (!Directory.Exists(destDirName))
                {
                    Directory.CreateDirectory(destDirName);
                    File.SetAttributes(destDirName, File.GetAttributes(sourceDirName));
                }

                if (destDirName[destDirName.Length - 1] != Path.DirectorySeparatorChar)
                    destDirName = destDirName + Path.DirectorySeparatorChar;

                string[] files = Directory.GetFiles(sourceDirName);
                foreach (string file in files)
                {
                    if (File.Exists(destDirName + Path.GetFileName(file)))
                        continue;
                    FileInfo fileInfo = new FileInfo(file);
                    if (fileInfo.Extension.Equals(".meta", StringComparison.CurrentCultureIgnoreCase))
                        continue;

                    File.Copy(file, destDirName + Path.GetFileName(file), true);
                    File.SetAttributes(destDirName + Path.GetFileName(file), FileAttributes.Normal);
                }

                string[] dirs = Directory.GetDirectories(sourceDirName);
                foreach (string dir in dirs)
                {
                    CopyDirectory(dir, destDirName + Path.GetFileName(dir));
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region GetFileBuffer 读取本地文件到byte数组
        /// <summary>
        /// 读取本地文件到byte数组
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns></returns>
        public static byte[] GetFileBuffer(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return null;
            }
            byte[] buffer = null;

            using (FileStream fs = new FileStream(filePath, FileMode.Open))
            {
                buffer = new byte[fs.Length];
                fs.Read(buffer, 0, buffer.Length);
            }
            return buffer;
        }
        #endregion
    }
}