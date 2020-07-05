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
using System.Text;

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

            if (!Directory.Exists(PathOut))
            {
                Directory.CreateDirectory(PathOut);
            }

            if (!Directory.Exists($"{PathOut}/Bytes"))
            {
                Directory.CreateDirectory($"{PathOut}/Bytes");
            }

            if (!Directory.Exists($"{PathOut}/Entities"))
            {
                Directory.CreateDirectory($"{PathOut}/Entities");
            }

            if (!Directory.Exists($"{PathOut}/Models"))
            {
                Directory.CreateDirectory($"{PathOut}/Models");
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
                CreateEnityFile(pathIn, sheet);
                CreateModelFile(pathIn, sheet);
            }

            File.Delete(pathTmp);
        }

        /// <summary>
        /// 生成 bytes 数据文件
        /// </summary>
        private void CreateData(string pathIn, ISheet sheet)
        {
            using (BinaryStream bs = new BinaryStream())
            {
                int rows = sheet.LastRowNum + 1;
                int cols = sheet.GetRow(sheet.LastRowNum).LastCellNum;

                // 计算表格实际行数
                int realRows = 0;
                for (int i = 0; i < rows; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue;
                    realRows++;
                }
                if (realRows < 3)
                {
                    Console.WriteLine($"{pathIn} 异常。");
                    return;
                }

                bs.WriteInt(realRows - 3);  //实际的数据条数
                bs.WriteInt(cols);

                string[,] tableHeadArr = new string[cols, 3];

                for (int i = 0; i < rows; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue;

                    for (int j = 0; j < cols; j++)
                    {
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
                string tableName = Path.GetFileNameWithoutExtension(pathIn);
                using (FileStream fs = new FileStream($"{PathOut}/Bytes/{tableName}.bytes", FileMode.Create))
                {
                    fs.Write(buffer, 0, buffer.Length);
                    Console.WriteLine($"生成 {tableName}.bytes 完毕 ...");
                }
            }
        }

        /// <summary>
        /// 生成 Entity 文件
        /// </summary>
        private void CreateEnityFile(string pathIn, ISheet sheet)
        {
            string template = TableTemplate.EntityTemplate;
            string tableName = Path.GetFileNameWithoutExtension(pathIn);

            // 替换时间与表名
            template = template.Replace("#CreateTime#", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            template = template.Replace("#TableName#", tableName);

            // 表格属性
            string properties = "";
            int cols = sheet.GetRow(sheet.LastRowNum).LastCellNum;
            for (int i = 0; i < cols; i++)
            {
                string propertyName = sheet.GetRow(0).GetCell(i).ToString();
                string propertyType = sheet.GetRow(1).GetCell(i).ToString();
                string propertyDesc = sheet.GetRow(2).GetCell(i).ToString();

                if (propertyName == "Id") continue;

                properties += $"        /// <summary>\n";
                properties += $"        /// {propertyDesc}\n";
                properties += $"        /// <summary>\n";
                properties += $"        public {propertyType} {propertyName} {{ get; set; }}";
                if (i != cols - 1)
                {
                    properties += "\n\n";
                }
            }
            template = template.Replace("#Properties#", properties);

            // 创建 Entity 文件
            using (FileStream fs = new FileStream($"{PathOut}/Entities/{tableName}Enity.cs", FileMode.Create))
            {
                byte[] buffer = Encoding.UTF8.GetBytes(template);
                fs.Write(buffer, 0, buffer.Length);
                Console.WriteLine($"生成 {tableName}Enity.cs 完毕 ...");
            }
        }

        /// <summary>
        /// 生成 Model 文件
        /// </summary>
        private void CreateModelFile(string pathIn, ISheet sheet)
        {
            string template = TableTemplate.ModelTemplate;
            string tableName = Path.GetFileNameWithoutExtension(pathIn);

            // 替换时间与表名
            template = template.Replace("#CreateTime#", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            template = template.Replace("#TableName#", tableName);

            // 表格属性
            string properties = "";
            int cols = sheet.GetRow(sheet.LastRowNum).LastCellNum;
            for (int i = 0; i < cols; i++)
            {
                string propertyName = sheet.GetRow(0).GetCell(i).ToString();
                string propertyType = sheet.GetRow(1).GetCell(i).ToString();

                properties += $"                entity.{propertyName} = {ConvertByte(propertyType)}bs.Read{MappingType(propertyType)}();\n";
            }
            template = template.Replace("#Properties#", properties);

            // 创建 Entity 文件
            using (FileStream fs = new FileStream($"{PathOut}/Models/{tableName}Model.cs", FileMode.Create))
            {
                byte[] buffer = Encoding.UTF8.GetBytes(template);
                fs.Write(buffer, 0, buffer.Length);
                Console.WriteLine($"生成 {tableName}Model.cs 完毕 ...");
            }
        }
        /// <summary>
        /// 类型映射
        /// </summary>
        private string MappingType(string type)
        {
            switch (type)
            {
                case "int":
                    return "Int";
                case "long":
                    return "Long";
                case "short":
                    return "Short";
                case "float":
                    return "Float";
                case "byte":
                    return "Byte";
                case "bool":
                    return "Bool";
                case "double":
                    return "Double";
                default:
                    return "UTF8String";
            }
        }

        /// <summary>
        /// 因为 MemoryStream 的 ReadByte 方法返回的是 int 类型，所以需要强转
        /// </summary>
        private object ConvertByte(string type)
        {
            return type == "byte" ? "(byte)" : "";
        }
    }
}
