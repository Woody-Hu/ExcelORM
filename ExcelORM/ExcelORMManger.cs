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

        /// <summary>
        /// xlsx文件名后缀标志
        /// </summary>
        private const string m_strxlsx = ".xlsx";

        /// <summary>
        /// xls文件名后缀标志
        /// </summary>
        private const string m_strxls = ".xls";
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

            TypeInfo useInfo = RegisteredType<T>();

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
            if (useFieInfo.Extension.ToLower().Equals(m_strxlsx))
            {
                useWorkBook = new XSSFWorkbook(useFieInfo.FullName);
            }
            else if (useFieInfo.Extension.ToLower().Equals(m_strxls))
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
       /// 尝试写出
       /// </summary>
       /// <typeparam name="T"></typeparam>
       /// <param name="inputPath">输入的路径</param>
       /// <param name="inputObjects">需写出的对象</param>
       /// <param name="overWriteIfExists">若存在是否删除</param>
       /// <returns></returns>
        public bool TryWrite<T>(string inputPath,List<T> inputObjects,bool overWriteIfExists = true)
        {
            //输入检查
            if (null == inputObjects || 0 == inputObjects.Count)
            {
                return false;
            }

            TypeInfo useInfo = RegisteredType<T>(true);

            //若注册失败
            if (null == useInfo)
            {
                return false;
            }

            FileInfo useFieInfo = new FileInfo(inputPath);

            //若文件不存在
            if (useFieInfo.Exists)
            {
                //若复写
                if (overWriteIfExists)
                {
                    try
                    {
                        //尝试删除
                        useFieInfo.Delete();
                    }
                    //异常保护
                    catch (Exception)
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }

            IWorkbook useWorkBook = null;

            //工厂制备WorkBook
            if (useFieInfo.Extension.ToLower().Equals(m_strxlsx))
            {
                useWorkBook = new XSSFWorkbook();
            }
            else if (useFieInfo.Extension.ToLower().Equals(m_strxls))
            {
                useWorkBook = new HSSFWorkbook();
            }

            //写出数据
            useInfo.WriteToWorkBook(useWorkBook, inputObjects.Cast<object>().ToList());

            try
            {
                using (Stream sw = new FileStream(inputPath, FileMode.CreateNew, FileAccess.Write))
                {
                    useWorkBook.Write(sw);
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
         
        }

        #region 私有字段
        /// <summary>
        /// 注册类型
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        private TypeInfo RegisteredType<T>(bool ifIsWrite = false)
        {
            Type useType = typeof(T);

            RegisteredType(useType, ifIsWrite);

            //进入读锁
            m_useReaderWriterLocker.EnterReadLock();
            var useInfo = m_useTypeMap[useType];
            //离开读锁
            m_useReaderWriterLocker.ExitReadLock();
            return useInfo;
        }

        /// <summary>
        /// 注册一个类
        /// </summary>
        /// <param name="inputType"></param>
        private void RegisteredType(Type inputType,bool ifIsWrite = false)
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
                WriteToDic(inputType,null);
                return;
            }

            //获取使用的类特性
            var useClassAtrribute = (ClassAttribute)classAttributes[0];

            //判断特性是否可用
            //若写状态则不限制
            if ((0 > useClassAtrribute.SheetIndex && string.IsNullOrWhiteSpace(useClassAtrribute.SheetName)) && !ifIsWrite)
            {
                WriteToDic(inputType, null);
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
                //且非写模式
                if ((null == tempPropertyAttribute ||
                    (0 > tempPropertyAttribute.UseColumnIndex
                    && string.IsNullOrWhiteSpace(tempPropertyAttribute.UseColumnName))) 
                    && !ifIsWrite)
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



            //注册
            if (0 != tempPropertyMap.Count)
            {
                WriteToDic(inputType, new TypeInfo(inputType, useClassAtrribute, tempPropertyMap));
            }
            else
            {
                WriteToDic(inputType, null);
            }


        }

        /// <summary>
        /// 写到字典
        /// </summary>
        /// <param name="inputType"></param>
        private static void WriteToDic(Type inputType,TypeInfo inputTypeInfo )
        {
            //进入写锁
            m_useReaderWriterLocker.EnterWriteLock();
            //内部检查
            if (!m_useTypeMap.ContainsKey(inputType))
            {
                m_useTypeMap.Add(inputType, inputTypeInfo);
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
            //进入读锁
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
        #endregion

    }
}
