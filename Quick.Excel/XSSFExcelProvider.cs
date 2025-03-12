using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Quick.Excel
{
    public class XSSFExcelProvider : ExcelProvider
    {
        protected override IWorkbook NewWorkbook() => new XSSFWorkbook();
    }
}
