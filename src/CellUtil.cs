//===================================================
//作    者：zhouzelong
//博    客：http://zhouzelong.cnblogs.com
//创建时间：2020-06-26 14:42:25
//备    注：
//===================================================

using NPOI.SS.UserModel;

namespace USGame
{
    /// <summary>
    /// ICell 扩展
    /// </summary>
    public static class CellUtil
    {
        public static int ToInt(this ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.Boolean:
                    return cell.BooleanCellValue ? 1 : 0;
                case CellType.String:
                    return int.Parse(cell.StringCellValue);
                case CellType.Numeric:
                    return (int)cell.NumericCellValue;
                default:
                    return 0;
            }
        }

        public static long ToLong(this ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.Boolean:
                    return cell.BooleanCellValue ? 1 : 0;
                case CellType.String:
                    return long.Parse(cell.StringCellValue);
                case CellType.Numeric:
                    return (long)cell.NumericCellValue;
                default:
                    return 0;
            }
        }

        public static short ToShort(this ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.Boolean:
                    return (short)(cell.BooleanCellValue ? 1 : 0);
                case CellType.String:
                    return short.Parse(cell.StringCellValue);
                case CellType.Numeric:
                    return (short)cell.NumericCellValue;
                default:
                    return 0;
            }
        }

        public static float ToFloat(this ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.Boolean:
                    return cell.BooleanCellValue ? 1f : 0f;
                case CellType.String:
                    return float.Parse(cell.StringCellValue);
                case CellType.Numeric:
                    return (float)cell.NumericCellValue;
                default:
                    return 0f;
            }
        }

        public static byte ToByte(this ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.Boolean:
                    return (byte)(cell.BooleanCellValue ? 1 : 0);
                case CellType.String:
                    return byte.Parse(cell.StringCellValue);
                case CellType.Numeric:
                    return (byte)cell.NumericCellValue;
                default:
                    return 0;
            }
        }

        public static bool ToBool(this ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                case CellType.String:
                    return bool.Parse(cell.StringCellValue);
                case CellType.Numeric:
                    return cell.NumericCellValue > 0d;
                default:
                    return false;
            }
        }

        public static double ToDouble(this ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.Boolean:
                    return cell.BooleanCellValue ? 1 : 0;
                case CellType.String:
                    return double.Parse(cell.StringCellValue);
                case CellType.Numeric:
                    return cell.NumericCellValue;
                default:
                    return 0d;
            }
        }

        public static string ToString(this ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Numeric:
                    return cell.NumericCellValue.ToString();
                default:
                    return "";
            }
        }
    }
}
