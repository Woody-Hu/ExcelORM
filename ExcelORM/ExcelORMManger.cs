using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelORM
{
    /// <summary>
    /// 对象映射管理器
    /// </summary>
    public class ExcelORMManger
    {
        #region 字段
        /// <summary>
        /// 使用的类特性
        /// </summary>
        private static Type m_useClassAttributeType = typeof(ClassAttribute);

        /// <summary>
        /// 使用的属性特性
        /// </summary>
        private static Type m_usePropertyAttributeType = typeof(PropertyAttribute);

        /// <summary>
        /// 使用的读写锁
        /// </summary>
        private static ReaderWriterLockSlim m_useReaderWriterLocker = new ReaderWriterLockSlim();

        /// <summary>
        /// 字典映射
        /// </summary>
        private static Dictionary<Type, TypeInfo> m_useTypeMap = new Dictionary<Type, TypeInfo>();

        /// <summary>
        /// 转换器字典
        /// </summary>
        private Dictionary<string, ChageValueDelegate> m_useChangeDelDic = new Dictionary<string, ChageValueDelegate>(); 
        #endregion

        /// <summary>
        /// 构造映射管理器
        /// </summary>
        /// <param name="inputAppendTransformer">输入的附加转换器字典</param>
        public ExcelORMManger(Dictionary<string, ChageValueDelegate> inputAppendTransformer = null)
        {
            if (null == inputAppendTransformer)
            {
                m_useChangeDelDic = new Dictionary<string, ChageValueDelegate>();
            }
            else
            {
                m_useChangeDelDic = inputAppendTransformer;
            }
        }

        /// <summary>
        /// 尝试读取
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="lstReadedValue"></param>
        /// <returns></returns>
        public bool TryRead<T>(string inputPath, out List<T> lstReadedValue)
            where T : class
        {
            lstReadedValue = new List<T>();

            Type useType = typeof(T);

            RegisteredType(useType);

            //进入读锁
            m_useReaderWriterLocker.EnterReadLock();
            var useInfo = m_useTypeMap[useType];
            //离开读锁
            m_useReaderWriterLocker.ExitReadLock();

            //若注册失败
            if (null == useInfo)
            {
                return false;
            }

            FileInfo useFieInfo = new FileInfo(inputPath);

            //若文件不存在
            if (!useFieInfo.Exists)
            {
                return false;
            }

            IWorkbook useWorkBook = null;

            //工厂制备WorkBook
            if (useFieInfo.Extension.ToLower().Equals(".xlsx"))
            {
                useWorkBook = new XSSFWorkbook(useFieInfo.FullName);
            }
            else if (useFieInfo.Extension.ToLower().Equals(".xls"))
            {
                using (FileStream fs = new FileStream(useFieInfo.FullName, FileMode.Open))
                {
                    useWorkBook = new HSSFWorkbook(fs);
                }

            }

            var returnValue = useInfo.ReadWorkBook(useWorkBook);

            lstReadedValue = returnValue.Cast<T>().ToList();

            return 0 != lstReadedValue.Count;


        }

        /// <summary>
        /// 注册一个类
        /// </summary>
        /// <param name="inputType"></param>
        private void RegisteredType(Type inputType)
        {
            //若已存在
            if (CheckInput(inputType))
            {
                return;
            }

            //获取类特性
            var classAttributes = inputType.GetCustomAttributes(m_useClassAttributeType, false);

            //获取检查
            if (null == classAttributes || 1 != classAttributes.Length)
            {
                //进入写锁
                m_useReaderWriterLocker.EnterWriteLock();
                m_useTypeMap.Add(inputType, null);
                //离开写锁
                m_useReaderWriterLocker.ExitWriteLock();
                return;
            }

            //获取使用的类特性
            var useClassAtrribute = (ClassAttribute)classAttributes[0];

            //判断特性是否可用
            if (0 > useClassAtrribute.SheetIndex && string.IsNullOrWhiteSpace(useClassAtrribute.SheetName))
            {
                //进入写锁
                m_useReaderWriterLocker.EnterWriteLock();
                m_useTypeMap.Add(inputType, null);
                //离开写锁
                m_useReaderWriterLocker.ExitWriteLock();
                return;
            }

            //临时局部变量
            PropertyAttribute tempPropertyAttribute = null;

            //属性-属性特性映射字典
            Dictionary<PropertyInfo, PropertyAttribute> tempPropertyMap
                = new Dictionary<PropertyInfo, PropertyAttribute>();

            //获取公开属性
            foreach (var oneProperty in inputType.GetProperties())
            {
                //不可读可写 跳过
                if (!oneProperty.CanRead || !oneProperty.CanWrite)
                {
                    continue;
                }

                //获取临时特性
                tempPropertyAttribute = oneProperty.GetCustomAttribute(m_usePropertyAttributeType) as PropertyAttribute;
                //没有特性跳过
                //特性属性检查
                if (null == tempPropertyAttribute ||
                    (0 > tempPropertyAttribute.UseColumnIndex
                    && string.IsNullOrWhiteSpace(tempPropertyAttribute.UseColumnName)))
                {
                    continue;
                }


                //若不是字符串类型且没有粘贴方法 且没有注册转换器
                if (!TypePropertyInfo.CheckProperty(oneProperty) && !CheckTransformer(tempPropertyAttribute))
                {
                    continue;
                }

                //若已注册转换器
                if (CheckTransformer(tempPropertyAttribute))
                {
                    //赋值转换器
                    tempPropertyAttribute.UseTransformer = m_useChangeDelDic[tempPropertyAttribute.UseTransformerName];
                }

                //添加到属性映射
                tempPropertyMap.Add(oneProperty, tempPropertyAttribute);
            }

            //进入写锁
            m_useReaderWriterLocker.EnterWriteLock();
            //注册
            if (0 != tempPropertyMap.Count)
            {
                m_useTypeMap.Add(inputType, new TypeInfo(inputType, useClassAtrribute, tempPropertyMap));
            }
            else
            {
                m_useTypeMap.Add(inputType, null);
            }
            //离开写锁
            m_useReaderWriterLocker.ExitWriteLock();
           
        }

        /// <summary>
        /// 检查输入
        /// </summary>
        /// <param name="inputType"></param>
        /// <returns></returns>
        private static bool CheckInput(Type inputType)
        {
            //进入读锁
            m_useReaderWriterLocker.EnterReadLock();
            bool returnValue = null == inputType || m_useTypeMap.ContainsKey(inputType);
            //进入写锁
            m_useReaderWriterLocker.ExitReadLock();
            return returnValue;
        }

        /// <summary>
        /// 检查转换器
        /// </summary>
        /// <param name="tempPropertyAttribute"></param>
        /// <returns></returns>
        private bool CheckTransformer(PropertyAttribute tempPropertyAttribute)
        {
            return (!string.IsNullOrWhiteSpace(tempPropertyAttribute.UseTransformerName) && m_useChangeDelDic.ContainsKey(tempPropertyAttribute.UseTransformerName) && null != m_useChangeDelDic[tempPropertyAttribute.UseTransformerName]);
        }

    }
}
