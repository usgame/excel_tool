//===================================================
//作    者：zhouzelong
//博    客：http://zhouzelong.cnblogs.com
//创建时间：2020-06-26 13:50:12
//备    注：
//===================================================

namespace USGame
{
    public static class TableTemplate
    {
        public static string EntityTemplate =
@"//===================================================
//作    者：zhouzelong
//博    客：http://zhouzelong.cnblogs.com
//创建时间：#CreateTime#
//备    注：
//===================================================

namespace USGame
{
    public class #TableName#Entity : TableEntity
    {
#Properties#
    }
}";

        public static string ModelTemplate =
@"//===================================================
//作    者：zhouzelong
//博    客：http://zhouzelong.cnblogs.com
//创建时间：#CreateTime#
//备    注：
//===================================================

namespace USGame
{
    public class #TableName#Model : TableModel<#TableName#Entity>
    {
        public override string TableName { get { return ""#TableName#""; } }

        public override void LoadList(BinaryStream bs)
        {
            int rows = bs.ReadInt();
            int columns = bs.ReadInt();

            for (int i = 0; i < rows; i++)
            {
                #TableName#Entity entity = new #TableName#Entity();
#Properties#
                m_List.Add(entity);
                m_Dic[entity.Id] = entity;
            }
        }
    }
}";
    }
}