using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelORM
{
    /// <summary>
    /// 类型属性封装
    /// </summary>
    internal class TypePropertyInfo
    {
        #region 私有字段
        /// <summary>
        /// 使用的属性封装
        /// </summary>
        private PropertyInfo m_usePropertyInfo = null;

        /// <summary>
        /// 使用的属性特性
        /// </summary>
        private PropertyAttribute m_usePropertyAttribute = null;

        /// <summary>
        /// 使用的属性类型
        /// </summary>
        private Type m_useProperType;

        /// <summary>
        /// 粘贴方法对象
        /// </summary>
        private MethodInfo m_useProperTypeParseMethod;

        /// <summary>
        /// 添加方法对象
        /// </summary>
        private MethodInfo m_useProperTypeAddMethod;

        /// <summary>
        /// 使用的列索引
        /// </summary>
        private int m_useColumnIndex;

        /// <summary>
        /// 当所属位置无值时是否使用上一个值
        /// </summary>
        private bool m_bUseLastValueWhenWhiteSpace = false;

        /// <summary>
        /// 读取完此字段对应的数据后是否切换到下一行（会开启无值时使用上一个值）
        /// </summary>
        private bool m_bChangeToNextRowWhenReadValue = false;

        /// <summary>
        /// 是否是Lst类型
        /// </summary>
        private bool m_bIfIsLstType = false;

        /// <summary>
        /// 使用的字符串类型
        /// </summary>
        private static Type m_useStringType = typeof(string);

        /// <summary>
        /// 使用的List泛型类型
        /// </summary>
        private static Type m_useListType = typeof(List<>);

        /// <summary>
        /// 粘贴方法名
        /// </summary>
        private const string m_useParseMethodName = "Parse";

        #endregion

        /// <summary>
        /// 构造方法
        /// </summary>
        /// <param name="inputPropertyInfo"></param>
        /// <param name="inputPropertyAttribute"></param>
        internal TypePropertyInfo(PropertyInfo inputPropertyInfo, PropertyAttribute inputPropertyAttribute)
        {
            m_usePropertyInfo = inputPropertyInfo;
            m_usePropertyAttribute = inputPropertyAttribute;
            m_useProperType = m_usePropertyInfo.PropertyType;
            m_bUseLastValueWhenWhiteSpace = m_usePropertyAttribute.UseLastValueWhenNull;
            m_bChangeToNextRowWhenReadValue = m_usePropertyAttribute.ChangeToNextRowWhenReadValue;

            //检查是否是列表类型
            IfIsLstType = CheckLstTypeProperty(m_useProperType);
        }

        /// <summary>
        /// 使用的列索引
        /// </summary>
        internal int UseColumnIndex
        {
            get { return m_useColumnIndex; }
            private set { m_useColumnIndex = value; }
        }

        /// <summary>
        /// 当所属位置无值时是否使用上一个值
        /// </summary>
        internal bool UseLastValueWhenWhiteSpace
        {
            get
            {
                return m_bUseLastValueWhenWhiteSpace;
            }

            private set
            {
                m_bUseLastValueWhenWhiteSpace = value;
            }
        }

        /// <summary>
        /// 是否是Lst类型
        /// </summary>
        internal bool IfIsLstType
        {
            get
            {
                return m_bIfIsLstType;
            }

            private set
            {
                m_bIfIsLstType = value;
            }
        }

        /// <summary>
        /// 读取完此字段对应的数据后是否切换到下一行（会开启无值时使用上一个值）
        /// </summary>
        internal bool ChangeToNextRowWhenReadValue
        {
            get
            {
                return m_bChangeToNextRowWhenReadValue;
            }

            private set
            {
                m_bChangeToNextRowWhenReadValue = value;
            }
        }

        /// <summary>
        /// 检查属性是否可用
        /// </summary>
        /// <param name="inputPropertyInfo"></param>
        /// <returns></returns>
        internal static bool CheckProperty(PropertyInfo inputPropertyInfo)
        {
            //获取属性类型
            var propertyType = inputPropertyInfo.PropertyType;
            return IfTypeCanUse(propertyType) || CheckLstTypeProperty(propertyType);
        }

        /// <summary>
        /// 准备数据
        /// </summary>
        /// <param name="inputSheet"></param>
        /// <param name="inputClassAttribute"></param>
        /// <param name="headerRowIndex"></param>
        internal void PrepareData(ISheet inputSheet, out int headerRowIndex)
        {
            headerRowIndex = 0;

            //已赋值列索引
            if (0 <= m_usePropertyAttribute.UseColumnIndex)
            {
                m_useColumnIndex = m_usePropertyAttribute.UseColumnIndex;
                return;
            }
            else
            {

                //行数保护
                var useHeaderLimitRowIndex = inputSheet.LastRowNum;

                for (int tempRowIndex = 0; tempRowIndex <= useHeaderLimitRowIndex; tempRowIndex++)
                {
                    //获取行
                    var tempRow = inputSheet.GetRow(tempRowIndex);

                    for (int tempColumnIndex = 0; tempColumnIndex <= tempRow.LastCellNum; tempColumnIndex++)
                    {
                        var tempCell = tempRow.GetCell(tempColumnIndex);

                        if (null == tempCell)
                        {
                            continue;
                        }

                        //若字符串匹配
                        if (tempCell.ToString().Equals(m_usePropertyAttribute.UseColumnName))
                        {
                            //列索引赋值
                            m_useColumnIndex = tempColumnIndex;
                            //使用行赋值(下一行）
                            headerRowIndex = tempRowIndex + 1;
                            return;
                        }
                    }
                }

            }


        }

        /// <summary>
        /// 设值
        /// </summary>
        /// <param name="inputObject"></param>
        /// <param name="inputValue"></param>
        internal void SetValue(object inputObject, string inputValue)
        {
            //单值模式
            if (!IfIsLstType)
            {
                SetSingleValue(inputObject, inputValue);
            }
            else
            {
                SetLstValue(inputObject, inputValue);
            }

        }

        #region 私有方法
        /// <summary>
        /// 使用的转换器
        /// </summary>
        private ChageValueDelegate ValueTransformer
        {
            get
            {
                return m_usePropertyAttribute.UseTransformer;
            }
        }

        /// <summary>
        /// 设置列表型数值
        /// </summary>
        /// <param name="inputObject"></param>
        /// <param name="inputValue"></param>
        private void SetLstValue(object inputObject, string inputValue)
        {
            object tempLstObj = null;
            tempLstObj = m_usePropertyInfo.GetValue(inputObject);

            //若属性没有被赋值
            if (null == tempLstObj)
            {
                //new一个0数量的list
                tempLstObj = Activator.CreateInstance(m_useProperType);
                //设值
                m_usePropertyInfo.SetValue(inputObject, tempLstObj);
            }

            //多态转换
            IList useLst = tempLstObj as IList;

            var useGenericType = m_useProperType.GetGenericArguments()[0];

            //若有转换器
            if (null != ValueTransformer)
            {
                //添加
                useLst.Add(ValueTransformer(inputValue));
            }
            else if (useGenericType == m_useStringType)
            {
                //添加
                useLst.Add(inputValue);
            }
            else
            {
                //数值转换
                var useValue = ChangeValue(inputValue, useGenericType);
                //添加
                useLst.Add(useValue);
            }
        }

        /// <summary>
        /// 设置单值数据
        /// </summary>
        /// <param name="inputObject"></param>
        /// <param name="inputValue"></param>
        private void SetSingleValue(object inputObject, string inputValue)
        {
            //若有转换器
            if (null != ValueTransformer)
            {
                //使用转换器
                m_usePropertyInfo.SetValue(inputObject, ValueTransformer(inputValue));
            }
            //若是字符串类型
            else if (m_useProperType == m_useStringType)
            {
                m_usePropertyInfo.SetValue(inputObject, inputValue);
            }
            else
            {
                object realValue = ChangeValue(inputValue, m_useProperType);
                //设值
                m_usePropertyInfo.SetValue(inputObject, realValue);
            }
        }

        /// <summary>
        /// 值类型转换
        /// </summary>
        /// <param name="inputValue"></param>
        /// <returns></returns>
        private object ChangeValue(string inputValue, Type inputType)
        {
            //设置粘贴方法引用
            if (null == m_useProperTypeParseMethod)
            {
                m_useProperTypeParseMethod = inputType.GetMethod(m_useParseMethodName, new Type[] { m_useStringType });
            }
            //转换
            var realValue = m_useProperTypeParseMethod.Invoke(null, new object[] { inputValue });
            return realValue;
        }

        /// <summary>
        /// 判断单一类型是否可以使用
        /// </summary>
        /// <param name="propertyType"></param>
        /// <returns></returns>
        private static bool IfTypeCanUse(Type propertyType)
        {
            //字符串类型 或 可 粘贴
            return m_useStringType == propertyType || IfTypeCanParse(propertyType);
        }

        /// <summary>
        /// 是否是可粘贴类型
        /// </summary>
        /// <param name="propertyType"></param>
        /// <returns></returns>
        private static bool IfTypeCanParse(Type propertyType)
        {
            return null != propertyType.GetMethod(m_useParseMethodName, new Type[] { m_useStringType });
        }

        /// <summary>
        /// 检查列表类型的属性
        /// </summary>
        /// <param name="inputPropertyInfo"></param>
        /// <returns></returns>
        private static bool CheckLstTypeProperty(Type inputPropertyType)
        {
            //是泛型 是 List<泛型> 泛型参数可粘贴
            return inputPropertyType.IsGenericType
                && inputPropertyType.GetGenericTypeDefinition() == m_useListType
                && IfTypeCanUse(inputPropertyType.GetGenericArguments()[0]);
        } 
        #endregion
    }
}
