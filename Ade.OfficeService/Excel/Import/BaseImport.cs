using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;

namespace Ade.OfficeService.Excel
{
    public abstract class BaseImport
    {
        private IWorkbook Workbook { get; set; }
        private ISheet Sheet { get; set; }

        protected virtual string SheetName
        {
            get
            {
                return sheetName;
            }
            set
            {
                sheetName = value;
            }
        }

        protected virtual int SheetIndex
        {
            get
            {
                return sheetIndex;
            }
            set
            {
                sheetIndex = value;
            }
        }

        /// <summary>
        /// 数据起始行
        /// </summary>
        protected virtual int StartRowIndex
        {
            get
            {
                return startRowIndex;
            }
            set
            {
                startRowIndex = value;
            }
        }

        /// <summary>
        /// 判断值在数据库中是否存在的委托。arg1:tableName,arg2:fieldName
        /// </summary>
        public virtual void SetDelegateDatabaseExist(Func<string, string, bool> delegateDatabaseExist)
        {
            this.delegateDatabaseExist = delegateDatabaseExist;
        }

        private int startRowIndex = 1;

        private int sheetIndex = 0;

        private string sheetName = "Sheet1";

        private Func<string, string, bool> delegateDatabaseExist;

        private ExcelHeaderRow HeaderRow { get; set; } = new ExcelHeaderRow();

        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="fileUrl"></param>
        private void Init(string fileUrl)
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

            if (SheetIndex != 0)
            {
                Sheet = Workbook.GetSheetAt(SheetIndex);
            }

            if (SheetName != "Sheet1")
            {
                Sheet = Workbook.GetSheet(SheetName);
            }

            if (Sheet == null)
            {
                Sheet = Workbook.GetSheetAt(0);
            }

            SetDictHeader();
        }

        /// <summary>
        /// 设置表头字典，列索引：列名
        /// </summary>
        /// <returns></returns>
        protected virtual void SetDictHeader()
        {
            IRow row = Sheet.GetRow(StartRowIndex - 1);
            ICell cell;
            for (int i = 0; i < row.PhysicalNumberOfCells; i++)
            {
                cell = row.GetCell(i);
                HeaderRow.Cells.Add(
                    new ExcelCol()
                    {
                        ColIndex = i,
                        ColName = cell?.StringCellValue ?? string.Empty
                    });
            }
        }

        /// <summary>
        /// 校验
        /// </summary>
        /// <returns></returns>
        public List<ExcelDataRow> Import(string fileUrl)
        {
            Init(fileUrl);

            List<ExcelDataRow> rows = ConvertISheet2DataRows(Sheet);
            AndFilter andFilter = new AndFilter() {  filters = GetFilters()};

            FilterContext context = new FilterContext()
            {
                DelegateDatabaseFilter = this.delegateDatabaseExist,
                TypeFilterInfo = TypeFilterInfoFactory.CreateInstance(this.GetType(), HeaderRow)
            };

            rows = andFilter.Filter(rows, context);

            return rows;
        }

        /// <summary>
        /// 获取所有的过滤器
        /// </summary>
        /// <returns></returns>
        private List<IFilter> GetFilters()
        {
            List<IFilter> filters = new List<IFilter>();
            List<BaseFilterAttribute> attrs = new List<BaseFilterAttribute>();
           TypeFilterInfo typeFilterInfo = TypeFilterInfoFactory.CreateInstance(this.GetType(), HeaderRow);

            typeFilterInfo.PropertyFilterInfos.ForEach(a => a.FilterAttrs.ForEach(f => attrs.Add(f)));

            attrs.Distinct(new FilterAttributeComparer()).ToList().ForEach
            (a =>
            {
                var filter = FilterFactory.CreateInstance(a.GetType());
                if (filter != null)
                {
                    filters.Add(filter);
                }
            });

            return filters;
        }


        /// <summary>
        /// 将ISheet转换为DataRows
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private List<ExcelDataRow> ConvertISheet2DataRows(ISheet sheet)
        {
            List<ExcelDataRow> rows = new List<ExcelDataRow>();

            for (int i = StartRowIndex; i < Sheet.PhysicalNumberOfRows; i++)
            {
                rows.Add(ConvertIRow2ExcelDataRow(sheet.GetRow(i)));
            }

            return rows;
        }

        /// <summary>
        /// 将IRow转换为ExcelDataRow
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private ExcelDataRow ConvertIRow2ExcelDataRow(IRow row)
        {
            ExcelDataRow dataRow = new ExcelDataRow()
            {
                IsValid = true,
                ErrorMsg = string.Empty,
                RowIndex = row.RowNum,
                DataCols = new List<ExcelDataCol>()
            };

            row.Cells.ForEach(c =>
            {
                dataRow.DataCols.Add(
                    new ExcelDataCol()
                    {
                        ColIndex = c.ColumnIndex,
                        ColValue = c.GetStringValue(),
                        RowIndex = dataRow.RowIndex,
                        ColName = HeaderRow?.Cells?.SingleOrDefault(h => h.ColIndex == c.ColumnIndex)?.ColName,
                    });
            });

            return dataRow;
        }
    }
}
