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
        #region 私有字段
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
        #endregion

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
        /// 读取
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        internal List<object> ReadWorkBook(IWorkbook input)
        {
            //准备数据
            PrepareDataForRead(input);

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
        /// 写出
        /// </summary>
        /// <param name="input"></param>
        /// <param name="lstInputValues"></param>
        internal void WriteToWorkBook(IWorkbook input,List<object> lstInputValues)
        {
            //准备数据
            PrepareDataForeWrite(input);

            int nowUseRowIndex = 1;
            //附加行号
            int appendIndexNumber = 0;
            Dictionary<int, IRow> useRowDic = new Dictionary<int, IRow>();

            foreach (var oneValue in lstInputValues)
            {
                //重置
                appendIndexNumber = 0;

                appendIndexNumber = WriteOneValue(nowUseRowIndex, appendIndexNumber, useRowDic, oneValue);

                nowUseRowIndex = nowUseRowIndex + appendIndexNumber;
            }
        }


        #region 私有方法
        /// <summary>
        /// 将一个对象写出
        /// </summary>
        /// <param name="nowUseRowIndex"></param>
        /// <param name="appendIndexNumber"></param>
        /// <param name="useRowDic"></param>
        /// <param name="oneValue"></param>
        /// <returns></returns>
        private int WriteOneValue(int nowUseRowIndex, int appendIndexNumber, Dictionary<int, IRow> useRowDic, object oneValue)
        {
            foreach (var onePropertyTypInfo in m_lstPropertyInfos)
            {
                //获取值
                var values = onePropertyTypInfo.GetValue(oneValue);
                //获取附加行号
                appendIndexNumber = Math.Max(appendIndexNumber, values.Count - 1);
                IRow tempRow;
                var tempIndex = nowUseRowIndex;
                //循环写值
                foreach (var onestrValue in values)
                {
                    if (!useRowDic.ContainsKey(tempIndex))
                    {
                        useRowDic.Add(tempIndex, m_useSheet.CreateRow(tempIndex));
                    }
                    tempRow = useRowDic[tempIndex];
                    //写值
                    var tempCell = tempRow.CreateCell(onePropertyTypInfo.UseColumnIndex);
                    tempCell.SetCellType(CellType.String);
                    tempCell.SetCellValue(onestrValue.ToString());
                    tempIndex++;
                }
            }

            return appendIndexNumber;
        }

        private void PrepareDataForeWrite(IWorkbook inputWorkbook)
        {
            m_useSheet = null;

            if (!string.IsNullOrEmpty(m_useClassAttribute.SheetName))
            {
                m_useSheet = inputWorkbook.CreateSheet(m_useClassAttribute.SheetName);
            }
            else
            {
                m_useSheet = inputWorkbook.CreateSheet();
            }

            //创建标头行
            IRow useHeaderRow = m_useSheet.CreateRow(0);

            int tempIndex = 0;
            HashSet<int> usedSet = new HashSet<int>();

            //初始化属性封装
            foreach (var onePropertyInfo in m_lstPropertyInfos)
            {
                //准备列数据
                onePropertyInfo.PrepareDataForWrite(useHeaderRow, ref tempIndex, ref usedSet);
            }
        }


        /// <summary>
        /// 数据准备
        /// </summary>
        /// <param name="inputWorkbook"></param>
        private void PrepareDataForRead(IWorkbook inputWorkbook)
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
                onePropertyInfo.PrepareDataForRead(m_useSheet, out tempDataRowIndex);
                //使用行冒泡
                useDataRowIndex = Math.Max(useDataRowIndex, tempDataRowIndex);
            }

            //设置使用数据起始行号
            m_dataStartRowNumber = this.m_useClassAttribute.RealUseDataStartRowIndex < 0 ? useDataRowIndex : this.m_useClassAttribute.RealUseDataStartRowIndex;

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
        private object SetValue(int useLength, string[] useValues, string[] readedRowValue
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
        private void ReadOneRow(int useLength, string[] useValues, string[] readedValues, ref int readedValue
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
        #endregion

    }
}
