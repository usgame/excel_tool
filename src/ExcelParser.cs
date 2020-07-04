//===================================================
//作    者：zhouzelong
//博    客：http://zhouzelong.cnblogs.com
//创建时间：2020-06-26 14:42:25
//备    注：
//===================================================

using System;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

namespace USGame
{
    public class ExcelParser : Singleton<ExcelParser>
    {
        /// <summary>
        /// 输出路径
        /// </summary>
        private string PathOut { get; set; }

        /// <summary>
        /// 单例，私有化构造函数
        /// </summary>
        private ExcelParser() { }

        /// <summary>
        /// 解析 Excel 表格
        /// </summary>
        /// <param name="pathIn"></param>
        /// <param name="pathOut"></param>
        public void Parse(string pathIn, string pathOut)
        {
            if (!string.IsNullOrEmpty(pathOut))
            {
                PathOut = Path.GetFullPath(pathOut);
            }
            else
            {
                PathOut = Path.GetFullPath("./");
            }

            pathIn = Path.GetFullPath(pathIn);
            bool isDirectory = Directory.Exists(pathIn);
            if (isDirectory)
            {
                ParseDirectory(pathIn);
            }
            else
            {
                ParseSingleExcel(pathIn);
            }
        }

        /// <summary>
        /// 解析目录
        /// </summary>
        private void ParseDirectory(string pathIn)
        {
            string[] paths = Directory.GetFiles(pathIn);
            for (int i = 0; i < paths.Length; i++)
            {
                ParseSingleExcel(paths[i]);
            }
        }

        /// <summary>
        /// 解析单个文件
        /// </summary>
        private void ParseSingleExcel(string pathIn)
        {
            if (!File.Exists(pathIn))
            {
                Console.WriteLine($"文件 {pathIn} 不存在");
                return;
            }

            bool isXls = Path.GetExtension(pathIn).Equals(".xls", StringComparison.CurrentCultureIgnoreCase);
            bool isXlsx = Path.GetExtension(pathIn).Equals(".xlsx", StringComparison.CurrentCultureIgnoreCase);
            if (!isXls && !isXlsx)
            {
                return;
            }

            string pathTmp = pathIn + ".tmp";
            if (File.Exists(pathTmp))
            {
                File.Delete(pathTmp);
            }
            File.Copy(pathIn, pathTmp);

            using (FileStream fs = new FileStream(pathTmp, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook;
                if (isXls)
                {
                    workbook = new HSSFWorkbook(fs);
                }
                else
                {
                    workbook = new XSSFWorkbook(fs);
                }
                ISheet sheet = workbook.GetSheetAt(0);
                CreateData(pathIn, sheet);
            }

            File.Delete(pathTmp);
        }

        /// <summary>
        /// 生成数据文件
        /// </summary>
        private void CreateData(string pathIn, ISheet sheet)
        {
            using (BinaryStream bs = new BinaryStream())
            {
                int rows = sheet.LastRowNum + 1;
                int cols = sheet.GetRow(sheet.LastRowNum).LastCellNum;

                string[,] tableHeadArr = new string[cols, 3];

                bs.WriteInt(rows - 3);
                bs.WriteInt(cols);

                for (int i = 0; i < rows; i++)
                {
                    IRow row = sheet.GetRow(i);
                    for (int j = 0; j < cols; j++)
                    {
                        if (row == null) continue;

                        ICell cell = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);

                        if (i < 3)
                        {
                            tableHeadArr[j, i] = cell.ToString();
                        }
                        else
                        {
                            string type = tableHeadArr[j, 1];
                            switch (type.ToLower())
                            {
                                case "int":
                                    bs.WriteInt(cell.ToInt());
                                    break;
                                case "long":
                                    bs.WriteLong(cell.ToLong());
                                    break;
                                case "short":
                                    bs.WriteShort(cell.ToShort());
                                    break;
                                case "float":
                                    bs.WriteFloat(cell.ToFloat());
                                    break;
                                case "byte":
                                    bs.WriteByte(cell.ToByte());
                                    break;
                                case "bool":
                                    bs.WriteBool(cell.ToBool());
                                    break;
                                case "double":
                                    bs.WriteDouble(cell.ToDouble());
                                    break;
                                default:
                                    bs.WriteUTF8String(cell.ToString());
                                    break;
                            }
                        }
                    }
                }

                // xor加密
                byte[] buffer = SecurityUtil.Xor(bs.ToArray());

                // zlib压缩
                buffer = ZlibHelper.CompressBytes(buffer);

                // 写入文件
                string fileName = Path.GetFileNameWithoutExtension(pathIn);
                using (FileStream fs = new FileStream($"{PathOut}/{fileName}.bytes", FileMode.Create))
                {
                    fs.Write(buffer, 0, buffer.Length);
                    Console.WriteLine($"生成 {fileName}.bytes 完毕 ...");
                }
            }
        }
    }
}
