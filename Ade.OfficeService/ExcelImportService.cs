using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using NPOI.SS.UserModel;

namespace Ade.OfficeService.Excel
{
    public static class ExcelImportService
    {
        private static IWorkbook Workbook { get; set; }
        private static ISheet Sheet { get; set; }

        private static ExcelHeaderRow HeaderRow { get; set; }

        /// <summary>
        /// 导入
        /// </summary>
        /// <typeparam name="T">模板类</typeparam>
        /// <param name="fileUrl">Excel文件绝对地址</param>
        /// <param name="delegateNotExistInDatabase">数据库校验委托，标记了数据库重复校验特性则必填</param>
        /// <returns></returns>
        public static List<ExcelDataRow> Validate<T>(string fileUrl, Func<DatabaseFilterContext, bool> delegateNotExistInDatabase = null)
            where T : class, new()
        {
            Init(fileUrl);

            List<ExcelDataRow> rows = ExcelConverter.Convert<T>(Sheet, HeaderRow, 1);
            AndFilter andFilter = new AndFilter() { filters = FiltersFlyWeight.CreateFilters<T>(HeaderRow) };

            FilterContext context = new FilterContext()
            {
                DelegateNotExistInDatabase = delegateNotExistInDatabase,
                TypeFilterInfo = TypeFilterInfoFlyweight.CreateInstance(typeof(T), HeaderRow)
            };

            rows = andFilter.Filter(rows, context);

            return rows;
        }

        /// <summary>
        /// 导出默认模板
        /// </summary>
        /// <typeparam name="T">模板类</typeparam>
        /// <returns></returns>
        public static IWorkbook ExportTemplate<T>()
            where T : class, new()
        {
            return ExcelExportService.Export(new List<T>() { });
        }

        public static IEnumerable<T> FastConvert<T>(List<ExcelDataRow> dataRows)
        {
            return dataRows.FastConvert<T>();
        }

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
    }
}
