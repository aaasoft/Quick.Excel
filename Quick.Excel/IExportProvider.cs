using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Quick.Excel
{
    public interface IExportProvider
    {
        /// <summary>
        /// 导出表格到指定的流
        /// </summary>
        /// <param name="table">表格对象</param>
        /// <param name="output">输出流</param>
        void ExportTable(table table, Stream output);
        /// <summary>
        /// 导出多个工作表的表格到指定的流
        /// </summary>
        /// <param name="tableDict"></param>
        /// <param name="output"></param>
        void ExportTable(Dictionary<string, table> tableDict, Stream output);
        /// <summary>
        /// 导出图片到指定的流
        /// </summary>
        /// <param name="title">标题</param>
        /// <param name="imageContent">图片内容</param>
        /// <param name="output">输出流</param>
        void ExportImage(string title, byte[] imageContent, Stream output);
    }
}
