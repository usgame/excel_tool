//===================================================
//作    者：zhouzelong
//博    客：http://zhouzelong.cnblogs.com
//创建时间：2020-06-26 13:50:28
//备    注：
//===================================================
using System.Collections.Generic;

namespace USGame
{
    /// <summary>
    /// 表格模型基类，管理一张表格
    /// </summary>
    /// <typeparam name="T">表格子类类型</typeparam>
    public abstract class TableModel<T> where T : TableEntity
    {
        /// <summary>
        /// 存储表格数据
        /// </summary>
        protected List<T> m_List;

        /// <summary>
        /// id与数据映射
        /// </summary>
        protected Dictionary<int, T> m_Dic;

        /// <summary>
        /// 构造函数
        /// </summary>
        public TableModel()
        {
            m_List = new List<T>();
            m_Dic = new Dictionary<int, T>();
        }

        /// <summary>
        /// 数据表名
        /// </summary>
        public abstract string TableName { get; }

        /// <summary>
        /// 加载数据列表
        /// </summary>
        public abstract void LoadList(BinaryStream bs);

        /// <summary>
        /// 加载数据表数据
        /// </summary>
        public void LoadData()
        {
            string tmppath = "./Bytes/" + TableName;
            string path = $"{System.IO.Path.GetFullPath(tmppath)}.bytes";
            byte[] buffer = IOUtil.GetFileBuffer(path);
            if (buffer == null)
            {
                System.Console.WriteLine($"{path} 不存在");
                return;
            }

            // 解压
            buffer = ZlibHelper.DecompressBytes(buffer);

            // 解密
            buffer = SecurityUtil.Xor(buffer);

            using (BinaryStream bs = new BinaryStream(buffer))
            {
                LoadList(bs);
            }
        }
    }
}