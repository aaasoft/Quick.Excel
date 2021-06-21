using System.Text;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.Collections.Generic;

namespace Quick.Excel
{
    public class ExcelProvider<T>
        where T : IWorkbook
    {
        /// <summary>
        /// 颜色序号数组
        /// </summary>
        public static readonly short[] FreeColorIndexArray = { 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 28, 29, 30, 31, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63 };

        private CellRangeAddress GetMergedRegion(ISheet sheet, int rowInd, int colInd)
        {
            for (var i = 0; i < sheet.NumMergedRegions; i++)
            {
                var mr = sheet.GetMergedRegion(i);
                if (mr.IsInRange(rowInd, colInd))
                    return mr;
            }
            return null;
        }

        private short GetColorIndex(IWorkbook workbook, byte[] color, ref int lastIndex)
        {
            if (workbook is NPOI.HSSF.UserModel.HSSFWorkbook)
            {
                var palette = ((NPOI.HSSF.UserModel.HSSFWorkbook)workbook).GetCustomPalette();
                var ret = palette.FindColor(color[0], color[1], color[2]);
                if (ret == null)
                {
                    palette.SetColorAtIndex(FreeColorIndexArray[lastIndex], color[0], color[1], color[2]);
                    var index = lastIndex;
                    lastIndex++;
                    return FreeColorIndexArray[index];
                }
                return ret.Indexed;
            }
            return -1;
        }

        public IWorkbook CreateWorkbook(table table)
        {
            return CreateWorkbook(new Dictionary<string, table>()
            {
                ["Sheet1"] = table
            });
        }

        public IWorkbook CreateWorkbook(Dictionary<string,table> tableDict)
        {
            IWorkbook workbook = System.Activator.CreateInstance<T>();
            foreach (var sheetName in tableDict.Keys)
            {
                var table = tableDict[sheetName];
                ISheet sheet = workbook.CreateSheet(sheetName);
                //当前行号
                var cRowInd = 0;
                //当前列号
                var cColInd = 0;
                //当前颜色序号
                var cColorIndex = 0;

                for (int i = 0; i < table.Count; i++)
                {
                    cColInd = 0;
                    while (sheet.IsMergedRegion(
                        new CellRangeAddress(cRowInd, cRowInd, cColInd, cColInd)))
                        cColInd++;

                    var tr = table[i];
                    IRow row = sheet.CreateRow(cRowInd);
                    for (int j = 0; j < tr.Count; j++)
                    {
                        while (sheet.IsMergedRegion(
                        new CellRangeAddress(cRowInd, cRowInd, cColInd, cColInd)))
                            cColInd++;

                        var td = tr[j];
                        ICell cell = row.CreateCell(cColInd);
                        //设置单元格的值
                        cell.SetCellValue(td.value);
                        //设置列的宽度自适应
                        var byteCount = string.IsNullOrEmpty(td.value) ? 1 : Encoding.Default.GetByteCount(td.value);
                        var needColumnWidth = byteCount * 256;
                        if (td.colspan > 1)
                        {
                            //如果跨列
                            var beforeColumnWidth = 0;
                            for (var z = 0; z < td.colspan; z++)
                            {
                                beforeColumnWidth += sheet.GetColumnWidth(cell.ColumnIndex + z);
                            }
                            var firstColumnWidth = sheet.GetColumnWidth(cell.ColumnIndex);
                            if (beforeColumnWidth < needColumnWidth)
                                sheet.SetColumnWidth(cell.ColumnIndex, firstColumnWidth + needColumnWidth - beforeColumnWidth);
                        }
                        else
                        {
                            //单个列
                            var beforeColumnWidth = sheet.GetColumnWidth(cell.ColumnIndex);
                            if (beforeColumnWidth < needColumnWidth)
                                sheet.SetColumnWidth(cell.ColumnIndex, needColumnWidth);
                        }

                        //合并单元格
                        if (td.rowspan > 1 || td.colspan > 1)
                        {
                            sheet.AddMergedRegion(new CellRangeAddress(
                                cRowInd,
                                cRowInd - 1 + td.rowspan,
                                cColInd,
                                cColInd - 1 + td.colspan));
                            if (td.colspan > 1)
                                cColInd += td.colspan - 1;
                        }

                        //单元格样式
                        ICellStyle cellStyle = workbook.CreateCellStyle();
                        //文字布局
                        switch (td.text_align)
                        {
                            case text_align.initial:
                                cellStyle.Alignment = HorizontalAlignment.General;
                                break;
                            case text_align.left:
                                cellStyle.Alignment = HorizontalAlignment.Left;
                                break;
                            case text_align.center:
                                cellStyle.Alignment = HorizontalAlignment.Center;
                                break;
                            case text_align.right:
                                cellStyle.Alignment = HorizontalAlignment.Right;
                                break;
                        }
                        cellStyle.VerticalAlignment = VerticalAlignment.Center;
                        //字体
                        IFont headerfont = workbook.CreateFont();
                        if (td.color_bytes != null)
                        {
                            if (workbook is NPOI.HSSF.UserModel.HSSFWorkbook)
                                headerfont.Color = GetColorIndex(workbook, td.color_bytes, ref cColorIndex);
                            else
                            {
                                var myFont = (NPOI.XSSF.UserModel.XSSFFont)headerfont;
                                myFont.SetColor(new NPOI.XSSF.UserModel.XSSFColor(td.color_bytes));
                            }
                        }
                        if (td.bold)
                            headerfont.IsBold = true;
                        cellStyle.SetFont(headerfont);
                        //背景
                        if (td.background_color_bytes != null)
                        {
                            if (workbook is NPOI.HSSF.UserModel.HSSFWorkbook)
                            {
                                cellStyle.FillForegroundColor = GetColorIndex(workbook, td.background_color_bytes, ref cColorIndex);
                                cellStyle.FillPattern = FillPattern.SolidForeground;
                            }
                            else
                            {
                                var myCellStyle = (NPOI.XSSF.UserModel.XSSFCellStyle)cellStyle;
                                myCellStyle.SetFillForegroundColor(new NPOI.XSSF.UserModel.XSSFColor(td.background_color_bytes));
                                cellStyle.FillPattern = FillPattern.SolidForeground;
                            }
                        }
                        //边框
                        if (td.border_color_bytes != null)
                        {
                            if (workbook is NPOI.HSSF.UserModel.HSSFWorkbook)
                            {
                                var color = GetColorIndex(workbook, td.border_color_bytes, ref cColorIndex);
                                cellStyle.BorderTop = BorderStyle.Thin;
                                cellStyle.TopBorderColor = color;
                                cellStyle.BorderRight = BorderStyle.Thin;
                                cellStyle.RightBorderColor = color;
                                cellStyle.BorderBottom = BorderStyle.Thin;
                                cellStyle.BottomBorderColor = color;
                                cellStyle.BorderLeft = BorderStyle.Thin;
                                cellStyle.LeftBorderColor = color;
                            }
                            else
                            {
                                var myCellStyle = (NPOI.XSSF.UserModel.XSSFCellStyle)cellStyle;
                                var color = new NPOI.XSSF.UserModel.XSSFColor(td.border_color_bytes);
                                myCellStyle.BorderTop = BorderStyle.Thin;
                                myCellStyle.SetTopBorderColor(color);
                                myCellStyle.BorderRight = BorderStyle.Thin;
                                myCellStyle.SetRightBorderColor(color);
                                myCellStyle.BorderBottom = BorderStyle.Thin;
                                myCellStyle.SetBottomBorderColor(color);
                                myCellStyle.BorderLeft = BorderStyle.Thin;
                                myCellStyle.SetLeftBorderColor(color);
                            }
                        }
                        //设置单元格样式
                        cell.CellStyle = cellStyle;

                        cColInd++;
                    }
                    cRowInd++;
                }
            }
            return workbook;
        }

        public void ExportTable(Dictionary<string, table> tableDict, Stream output)
        {
            var workbook = CreateWorkbook(tableDict);
            workbook.Write(output);
            workbook.Close();
            workbook = null;
        }

        public void ExportTable(table table, Stream output)
        {
            var workbook = CreateWorkbook(table);
            workbook.Write(output);
            workbook.Close();
            workbook = null;
        }

        public IWorkbook CreateWorkbook(string title, byte[] imageContent)
        {
            IWorkbook workbook = System.Activator.CreateInstance<T>();
            ISheet sheet = workbook.CreateSheet("Sheet1");

            //添加图片
            int pictureIdx = workbook.AddPicture(imageContent, PictureType.PNG);
            IDrawing patriarch = sheet.CreateDrawingPatriarch();
            //放图片的位置
            var creationHelper = workbook.GetCreationHelper();
            IClientAnchor anchor = creationHelper.CreateClientAnchor();
            anchor.AnchorType = AnchorType.MoveAndResize;
            anchor.Col1 = 0;
            anchor.Row1 = 1;
            anchor.Col2 = 0;
            anchor.Row2 = 1;
            //如果未设置了标题
            if (string.IsNullOrEmpty(title))
            {
                anchor.Row1 = 0;
                anchor.Row2 = 0;
            }
            IPicture pict = patriarch.CreatePicture(anchor, pictureIdx);
            pict.Resize();

            //如果设置了标题
            if (!string.IsNullOrEmpty(title))
            {
                ICellStyle HeadercellStyle = workbook.CreateCellStyle();
                HeadercellStyle.Alignment = HorizontalAlignment.Center;
                HeadercellStyle.VerticalAlignment = VerticalAlignment.Center;

                //字体
                IFont headerfont = workbook.CreateFont();
                headerfont.IsBold = true;
                HeadercellStyle.SetFont(headerfont);

                var row_0 = sheet.CreateRow(0);
                var cell_0 = row_0.CreateCell(0);
                cell_0.CellStyle = HeadercellStyle;
                cell_0.SetCellValue(title);

                sheet.AddMergedRegion(new CellRangeAddress(
                            0,
                            0,
                            pict.ClientAnchor.Col1,
                            pict.ClientAnchor.Col2 - 1));
            }
            return workbook;
        }

        public void ExportImage(string title, byte[] imageContent, Stream output)
        {
            var workbook = CreateWorkbook(title, imageContent);
            workbook.Write(output);
            workbook = null;
        }
    }
}
