using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelORM
{
    /// <summary>
    /// 字符串转换为期望对象委托
    /// </summary>
    /// <param name="input"></param>
    /// <returns></returns>
    public delegate object ChageValueDelegate(string input);
}
