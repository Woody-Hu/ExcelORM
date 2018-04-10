using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelORM
{
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
    /// <summary>
    /// 类特性
    /// </summary>
    public class ClassAttribute : Attribute
    {
        /// <summary>
        /// 使用的表索引
        /// </summary>
        private int m_sheetIndex = -1;

        /// <summary>
        /// 使用的表名
        /// </summary>
        private string m_sheetName = null;

        /// <summary>
        /// 使用的数据起始行索引
        /// </summary>
        private int m_realUseDataStartRowIndex = -1;

        /// <summary>
        /// 使用的数据起始行索引
        /// </summary>
        public int RealUseDataStartRowIndex
        {
            get { return m_realUseDataStartRowIndex; }
            set { m_realUseDataStartRowIndex = value; }
        }

        /// <summary>
        /// 使用的表索引
        /// </summary>
        public int SheetIndex
        {
            get
            {
                return m_sheetIndex;
            }

            set
            {
                m_sheetIndex = value;
            }
        }

        /// <summary>
        /// 使用的表名
        /// </summary>
        public string SheetName
        {
            get
            {
                return m_sheetName;
            }

            set
            {
                m_sheetName = value;
            }
        }
    }
}
