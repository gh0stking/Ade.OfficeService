using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Ade.OfficeService.Excel
{
    public static class ExcelConverter
    {
        private static Hashtable Table = Hashtable.Synchronized(new Hashtable(1024));

        /// <summary>
        /// 将Sheet转化为ExcelDataRow集合
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startDataRowIndex"></param>
        /// <returns></returns>
        public static List<ExcelDataRow> Convert<TTemplate>(ISheet sheet, ExcelHeaderRow headerRow, int startDataRowIndex)
        {
            List<ExcelDataRow> dataRows = new List<ExcelDataRow>();
            for (int i = startDataRowIndex; i < sheet.PhysicalNumberOfRows; i++)
            {
                dataRows.Add(Convert<TTemplate>(sheet.GetRow(i), headerRow));
            }

            return dataRows;
        }

        /// <summary>
        /// 将IRow转换为ExcelDataRow
        /// </summary>
        /// <typeparam name="TTemplate"></typeparam>
        /// <param name="row"></param>
        /// <param name="headerRow"></param>
        /// <returns></returns>
        public static ExcelDataRow Convert<TTemplate>(IRow row, ExcelHeaderRow headerRow)
        {
            Type type = typeof(TTemplate);

            ExcelDataRow dataRow = new ExcelDataRow()
            {
                DataCols = new List<ExcelDataCol>(),
                ErrorMsg = string.Empty,
                IsValid = true,
                RowIndex = row.RowNum
            };

            ICell cell;
            ExcelDataCol dataCol;
            string colName;
            string propertyName;
            string key;
            for (int i = 0; i < headerRow.Cells.Count; i++)
            {
                colName = headerRow?.Cells?.SingleOrDefault(h => h.ColIndex == i)?.ColName;

                if (colName == null)
                {
                    continue;
                }

                key = $"{type.FullName}_{i}";

                if (Table[key] == null)
                {
                    propertyName = type.GetProperties().ToList().FirstOrDefault(p => p.IsDefined(typeof(ExcelImportAttribute), false)
                && p.GetCustomAttribute<ExcelImportAttribute>()?.ColName == colName
                )?.Name;
                    Table[key] = propertyName;
                }

                dataCol = new ExcelDataCol()
                {
                    ColIndex = i,
                    ColName = colName,
                    PropertyName = Table[key]?.ToString(),
                    RowIndex = row.RowNum,
                    ColValue = row.GetCell(i) == null ? string.Empty : row.GetCell(i).GetStringValue()
                };

                dataRow.DataCols.Add(dataCol);
            }

            return dataRow;
        }
    }
}
