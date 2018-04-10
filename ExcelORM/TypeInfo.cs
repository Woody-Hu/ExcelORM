using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelORM
{
    /// <summary>
    /// 使用的类型信息封装
    /// </summary>
    internal class TypeInfo
    {
        /// <summary>
        /// 使用的Type
        /// </summary>
        private Type m_thisType;

        /// <summary>
        /// 使用的类特性
        /// </summary>
        private ClassAttribute m_useClassAttribute;

        /// <summary>
        /// 数据起始行
        /// </summary>
        private int m_dataStartRowNumber = 0;

        /// <summary>
        /// 使用的属性封装
        /// </summary>
        private List<TypePropertyInfo> m_lstPropertyInfos = new List<TypePropertyInfo>();

        /// <summary>
        /// 使用的表对象
        /// </summary>
        private ISheet m_useSheet = null;

        /// <summary>
        /// 构造方法
        /// </summary>
        /// <param name="useType"></param>
        /// <param name="useClassAttribute"></param>
        internal TypeInfo(Type useType, ClassAttribute useClassAttribute, Dictionary<PropertyInfo, PropertyAttribute> inputPropertyMap)
        {
            m_thisType = useType;
            m_useClassAttribute = useClassAttribute;
            //制备成员
            foreach (var oneKVP in inputPropertyMap)
            {
                m_lstPropertyInfos.Add(new TypePropertyInfo(oneKVP.Key, oneKVP.Value));
            }

        }

        /// <summary>
        /// 数据准备
        /// </summary>
        /// <param name="inputWorkbook"></param>
        private void PrepareData(IWorkbook inputWorkbook)
        {
            m_useSheet = null;
            m_dataStartRowNumber = 0;

            //利用索引
            if (0 <= m_useClassAttribute.SheetIndex)
            {
                m_useSheet = inputWorkbook.GetSheetAt(m_useClassAttribute.SheetIndex);
            }
            else
            {
                m_useSheet = inputWorkbook.GetSheet(m_useClassAttribute.SheetName);
            }

            int useDataRowIndex = 0;

            int tempDataRowIndex = 0;

            //初始化属性封装
            foreach (var onePropertyInfo in m_lstPropertyInfos)
            {
                onePropertyInfo.PrepareData(m_useSheet, out tempDataRowIndex);
                //使用行冒泡
                useDataRowIndex = Math.Max(useDataRowIndex, tempDataRowIndex);
            }

            //设置使用数据起始行号
            m_dataStartRowNumber = this.m_useClassAttribute.RealUseDataStartRowIndex < 0 ? useDataRowIndex : this.m_useClassAttribute.RealUseDataStartRowIndex;

        }

        /// <summary>
        /// 读取
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        internal List<object> ReadWorkBook(IWorkbook input)
        {
            //准备数据
            PrepareData(input);

            int useLength = m_lstPropertyInfos.Count;

            string[] useValues = new string[useLength];
            string[] readRowValues = new string[useLength];

            List<object> returnValues = new List<object>();

            //上次的生成对象
            object lastValue = null;

            //获取的非列表型属性数量
            int nowGetNoneLstPropertyValue = 0;
            //读取到的值
            int readedValue = 0;

            //跳过标示
            bool continueTag = false;

            //逐行读取
            for (int useRowIndex = m_dataStartRowNumber; useRowIndex <= m_useSheet.LastRowNum; useRowIndex++)
            {
                //重置
                nowGetNoneLstPropertyValue = 0;
                continueTag = false;
                readedValue = 0;

                var useRow = m_useSheet.GetRow(useRowIndex);

                //全Null行
                if (null == useRow)
                {
                    continue;
                }

                //读取一行
                ReadOneRow(useLength, useValues, readRowValues,
                    ref readedValue, ref nowGetNoneLstPropertyValue, ref continueTag, useRow);

                //若需跳转
                if (continueTag && 1 == readedValue)
                {
                    continue;
                }

                try
                {
                    //跳过空数据组
                    if (useLength == (from n in useValues where null == n select n).Count())
                    {
                        continue;
                    }

                    lastValue = SetValue(useLength, useValues,readRowValues, returnValues, lastValue, nowGetNoneLstPropertyValue);

                }
                //异常跳过
                catch (Exception)
                {
                    continue;
                }

            }

            return returnValues;
        }

        /// <summary>
        /// 设置值
        /// </summary>
        /// <param name="useLength"></param>
        /// <param name="useValues"></param>
        /// <param name="returnValues"></param>
        /// <param name="lastValue"></param>
        /// <param name="nowGetNoneLstPropertyValue"></param>
        /// <returns></returns>
        private object SetValue(int useLength, string[] useValues,string[] readedRowValue
            , List<object> returnValues, object lastValue, int nowGetNoneLstPropertyValue)
        {
            //使用的临时对象
            object tempObject;

            //是否创建标示
            bool ifCreat = false;

            //若获取到非列表型数据 或没有上次数据
            if (nowGetNoneLstPropertyValue > 0 || null == lastValue)
            {
                //创建对象
                tempObject = Activator.CreateInstance(m_thisType);
                //添加到返回列表
                returnValues.Add(tempObject);
                ifCreat = true;
            }
            else
            {
                //使用上次的对象
                tempObject = lastValue;
            }


            //属性设值
            for (int propertyIndex = 0; propertyIndex < useLength; propertyIndex++)
            {
                string useValue;
                //若创建且是列表型属性
                if (ifCreat && m_lstPropertyInfos[propertyIndex].IfIsLstType)
                {
                    useValue = readedRowValue[propertyIndex];
                    //重置记录结果
                    useValues[propertyIndex] = useValue;
                }
                else
                {
                    useValue = useValues[propertyIndex];
                }

                if (null != useValue)
                {
                    m_lstPropertyInfos[propertyIndex].SetValue(tempObject, useValue);
                }
            }

            //保存当次的值
            lastValue = tempObject;
            return lastValue;
        }

        /// <summary>
        /// 读取一行
        /// </summary>
        /// <param name="useLength"></param>
        /// <param name="useValues"></param>
        /// <param name="nowGetNoneLstPropertyValue"></param>
        /// <param name="continueTag"></param>
        /// <param name="useRow"></param>
        /// <param name="readedValue"></param>
        private void ReadOneRow(int useLength, string[] useValues, string[] readedValues,ref int readedValue
            , ref int nowGetNoneLstPropertyValue, ref bool continueTag, IRow useRow)
        {
            //读取列
            for (int propertyIndex = 0; propertyIndex < useLength; propertyIndex++)
            {
                var tempProperty = m_lstPropertyInfos[propertyIndex];
               
                //重置临时值
                readedValues[propertyIndex] = null;

                //若不使用上次值
                if (!tempProperty.UseLastValueWhenWhiteSpace && !tempProperty.ChangeToNextRowWhenReadValue)
                {
                    //值重置
                    useValues[propertyIndex] = null;
                }

                var useCell = useRow.GetCell(tempProperty.UseColumnIndex);

                if (null == useCell || string.IsNullOrWhiteSpace(useCell.ToString()))
                {
                    continue;
                }

                //若是非列表属性
                if (!tempProperty.IfIsLstType)
                {
                    nowGetNoneLstPropertyValue++;
                }

                useValues[propertyIndex] = useCell.ToString();
                readedValues[propertyIndex] = useCell.ToString();

                //若读取后跳转
                if (tempProperty.ChangeToNextRowWhenReadValue)
                {
                    continueTag = true;
                }

                //添加已读取属性
                readedValue++;
            }
        }


    }
}
