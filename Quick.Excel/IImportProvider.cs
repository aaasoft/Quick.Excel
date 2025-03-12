using System;
using System.Collections.Generic;
using System.IO;

namespace Quick.Excel
{
    public interface IImportProvider
    {
        /// <summary>
        /// 导入表格
        /// </summary>
        /// <param name="output"></param>
        /// <returns></returns>
        Dictionary<string, table> ImportTable(Stream output);
    }
}