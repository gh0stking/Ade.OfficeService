using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using NPOI.SS.UserModel;

namespace Ade.OfficeService.Excel
{
    public static class ExcelImportService<T>
        where T:IExcelImport
    {
        private static IWorkbook Workbook { get; set; }
        private static ISheet Sheet { get; set; }

        private static ExcelHeaderRow HeaderRow { get; set; }

        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="fileUrl"></param>
        private static void Init(string fileUrl)
        {
            if (string.IsNullOrWhiteSpace(fileUrl))
            {
                throw new ArgumentNullException("fileUrl");
            }

            if (!File.Exists(fileUrl))
            {
                throw new FileNotFoundException();
            }

            string ext = Path.GetExtension(fileUrl).ToLower().Trim();
            if (ext != ".xls" && ext != ".xlsx")
            {
                throw new NotSupportedException("非法文件");
            }

            try
            {
                Workbook = WorkbookFactory.Create(fileUrl);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            Sheet = Workbook.GetSheetAt(0);

            SetDictHeader();
        }

        /// <summary>
        /// 设置表头字典，列索引：列名
        /// </summary>
        /// <returns></returns>
        private static void SetDictHeader()
        {
            HeaderRow = new ExcelHeaderRow();
            IRow row = Sheet.GetRow(0);
            ICell cell;
            for (int i = 0; i < row.PhysicalNumberOfCells; i++)
            {
                cell = row.GetCell(i);
                HeaderRow.Cells.Add(
                    new ExcelCol()
                    {
                        ColIndex = i,
                        ColName = cell.GetStringValue()
                    });
            }
        }

        /// <summary>
        /// 校验
        /// </summary>
        /// <returns></returns>
        public static List<ExcelDataRow> Import(string fileUrl, Func<DatabaseFilterContext, bool> func = null)
        {
            Init(fileUrl);

            List<ExcelDataRow> rows = ExcelConverter.Convert<T>(Sheet, HeaderRow, 1);
            AndFilter andFilter = new AndFilter() { filters = FiltersFactory.CreateFilters<T>(HeaderRow) };

            FilterContext context = new FilterContext()
            {
                DelegateDatabaseFilter = func,
                TypeFilterInfo = TypeFilterInfoFactory.CreateInstance(typeof(T), HeaderRow)
            };

            rows = andFilter.Filter(rows, context);

            return rows;
        }
    }
}
