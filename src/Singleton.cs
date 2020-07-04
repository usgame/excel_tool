//===================================================
//作    者：zhouzelong
//博    客：http://zhouzelong.cnblogs.com
//创建时间：2020-06-26 13:44:52
//备    注：
//===================================================
using System;

namespace USGame
{
    // 由于单例基类不能实例化，故设计为抽象类
    public abstract class Singleton<T> where T : class
    {
        class Nested
        {
            // 创建模板类实例，参数2设为true表示支持私有构造函数
            internal static readonly T instance = Activator.CreateInstance(typeof(T), true) as T;
        }
        public static T Instance { get { return Nested.instance; } }
    }
}