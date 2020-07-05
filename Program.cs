//===================================================
//作    者：zhouzelong
//博    客：http://zhouzelong.cnblogs.com
//创建时间：2020-06-26 14:42:25
//备    注：
//===================================================
using System;

namespace USGame
{
    class Program
    {
        /// <summary>
        /// 程序入口
        /// </summary>
        /// <param name="args">传入参数</param>
        static void Main(string[] args)
        {
            // 参数是否正确
            if (args.Length == 0)
            {
                Console.WriteLine("参数不正确\n使用格式：excel_tool 文件(夹)路径 [输出文件夹路径]");
                Pause();
                return;
            }

            // 解析 excel 文件
            string pathIn = args[0];
            string pathOut = null;
            if (args.Length > 1)
            {
                pathOut = args[1];
            }

            try
            {
                ExcelParser.Instance.Parse(pathIn, pathOut);
            }
            catch
            {
            }
            finally
            {
                Pause();
            }
        }

        /// <summary>
        /// 程序暂停
        /// </summary>
        static void Pause()
        {
            Console.WriteLine("按任意键退出窗口...");
            Console.ReadKey();
        }
    }
}
