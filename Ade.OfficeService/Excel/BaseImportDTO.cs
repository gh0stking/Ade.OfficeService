using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using NPOI.SS.UserModel;

namespace Ade.OfficeService.Excel
{
    public class BaseImportDTO
    {
        public IWorkbook Workbook { get; set; }
        public ISheet Sheet { get; set; }

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

        protected virtual Func<string, string, bool> DBExistDelegate
        {
            get
            {
                return dbExistDelegate;
            }
            set
            {
                dbExistDelegate = value;
            }
        }

        private int startRowIndex = 1;

        private int sheetIndex = 0;

        private string sheetName = "Sheet1";

        private Func<string, string, bool> dbExistDelegate;

        private List<ExcelCol> HeaderRow { get; set; } = new List<ExcelCol>();
        public void Init(string fileUrl)
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
        }

        /// <summary>
        /// 设置表头字典，列索引：列名
        /// </summary>
        /// <returns></returns>
        protected virtual void SetDictHeader()
        {
            IRow row = Sheet.GetRow(StartRowIndex - 1);
            ICell cell;
            for (int i = 0; i <= row.PhysicalNumberOfCells; i++)
            {
                cell = row.GetCell(i);
                HeaderRow.Add(new ExcelCol() { ColIndex = i, ColName = cell.StringCellValue });
            }
        }

        protected virtual List<RowWrapper> Validate(IWorkbook workbook)
        {
            List<RowWrapper> rowWrappers = new List<RowWrapper>();

            for (int i = StartRowIndex; i < Sheet.PhysicalNumberOfRows; i++)
            {
                rowWrappers.Add(new RowWrapper()
                {
                    Row = Sheet.GetRow(i),
                    RowIndex = i,
                    IsValid = true,
                    ErrorMsg = string.Empty
                });
            }


        }
    }
}
