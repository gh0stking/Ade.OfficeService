//using NPOI.HSSF.UserModel;
//using NPOI.HSSF.Util;
//using NPOI.SS.UserModel;
//using NPOI.SS.Util;
//using System.Collections.Generic;
//using System.IO;
//using System.Linq;

//namespace Ade.OfficeService.Excel
//{
//    public class BaseExport
//    {
//        private IWorkbook Workbook { get; set; }
//        private ISheet Sheet { get; set; }
       
//        public void Init()
//        {
//            Workbook = new HSSFWorkbook();
//            Sheet = Workbook.CreateSheet();
//        }

//        public IWorkbook Export(List<BaseExport> baseExportDtos)
//        {
//            ExcelSetHeader();

//            ExcelSetDataRow(baseExportDtos);

//            return this.Workbook;
//        }

//        /// <summary>
//        /// 设置表头
//        /// </summary>
//        protected virtual void ExcelSetHeader()
//        {
//            IRow row = Sheet.CreateRow(0);
//            int colIndex = 0;
//            foreach (var item in ExportInfoWrapper.MappingDict)
//            {
//                row.CreateCell(colIndex).SetCellValue(item.Value);
//                colIndex++;
//            }
//        }

//        /// <summary>
//        /// 设置数据
//        /// </summary>
//        /// <param name="lst"></param>
//        protected virtual void ExcelSetDataRow(List<BaseExport> lst)
//        {
//            IRow row;
//            for (int i = 0; i < lst.Count; i++)
//            {
//                row = Sheet.CreateRow(i + 1);
//                int colIndex = 0;
//                BaseExport dto = lst[i];
//                foreach (var item in ExportInfoWrapper.MappingDict)
//                {
//                    //全部以字符串输出,格式化在准备数据时处理好
//                    row.CreateCell(colIndex).SetCellValue(ExportHelper.GetValue(dto, item.Key));
//                    colIndex++;
//                }
//            }
//        }
//    }
//}
