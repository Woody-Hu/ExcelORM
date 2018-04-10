using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelORM
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    /// <summary>
    /// 属性特性
    /// </summary>
    public class PropertyAttribute : Attribute
    {
        #region 使用的转换器名称
        /// <summary>
        /// 使用的列索引
        /// </summary>
        private int m_useColumnIndex = -1;

        /// <summary>
        /// 使用的列名称
        /// </summary>
        private string m_useColumnName = null;

        /// <summary>
        /// 当所属位置无值时是否使用上一个值
        /// </summary>
        private bool m_bUseLastValueWhenNull = false;

        /// <summary>
        /// 读取完此字段对应的数据后是否切换到下一行（会开启无值时使用上一个值）
        /// </summary>
        private bool m_bChangeToNextRowWhenReadValue = false;

        /// <summary>
        /// 使用的转换器名称
        /// </summary>
        private string m_strUseTransformerName = null;

        /// <summary>
        /// 使用的转换器
        /// </summary>
        private ChageValueDelegate m_useTransformer = null;
        #endregion

        /// <summary>
        /// 使用的列索引
        /// </summary>
        public int UseColumnIndex
        {
            get
            {
                return m_useColumnIndex;
            }

            set
            {
                m_useColumnIndex = value;
            }
        }

        /// <summary>
        /// 使用的列名称
        /// </summary>
        public string UseColumnName
        {
            get
            {
                return m_useColumnName;
            }

            set
            {
                m_useColumnName = value;
            }
        }

        /// <summary>
        /// 当所属位置无值时是否使用上一个值
        /// </summary>
        public bool UseLastValueWhenNull
        {
            get
            {
                return m_bUseLastValueWhenNull;
            }

            set
            {
                m_bUseLastValueWhenNull = value;
            }
        }

        /// <summary>
        /// 读取完此字段对应的数据后是否切换到下一行（会开启无值时使用上一个值）
        /// </summary>
        public bool ChangeToNextRowWhenReadValue
        {
            get
            {
                return m_bChangeToNextRowWhenReadValue;
            }

            set
            {
                m_bChangeToNextRowWhenReadValue = value;
            }
        }

        /// <summary>
        /// 使用的转换器名称
        /// </summary>
        public string UseTransformerName
        {
            get
            {
                return m_strUseTransformerName;
            }

            set
            {
                m_strUseTransformerName = value;
            }
        }

        /// <summary>
        /// 使用的转换器
        /// </summary>
        internal ChageValueDelegate UseTransformer
        {
            get
            {
                return m_useTransformer;
            }

            set
            {
                m_useTransformer = value;
            }
        }
    }
}
