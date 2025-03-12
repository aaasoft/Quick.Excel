using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace Quick.Excel
{
    public class HSSFExcelProvider : ExcelProvider
    {
        protected override IWorkbook NewWorkbook() => new HSSFWorkbook();
        protected override IWorkbook OpenWorkbook(Stream stream) => new HSSFWorkbook(stream);
    }
}
