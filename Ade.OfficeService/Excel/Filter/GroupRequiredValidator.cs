using QDMice.Api.Impl.Repositories;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ade.OfficeService.Excel
{
    [FilterOption(Order = 8)]
    public class ValidateGroupRequired : BaseImportValidator
    {
        public Dictionary<string, List<string>> DictGroupRequired { get; set; } = new Dictionary<string, List<string>>();

        public override void OnColumnLoopEnd(string cellStringValue, ValidateAttrsWrapper validateAttrsWrapper, ImportRowWrapper myRow, string excelName, IMice_ExcelImportRepository mice_ExcelImportRepository)
        {
            if (validateAttrsWrapper.ValidateGroupRequired)
            {
                if (!DictGroupRequired.ContainsKey(validateAttrsWrapper.RequiredGroupName))
                {
                    DictGroupRequired[validateAttrsWrapper.RequiredGroupName] = new List<string>();
                }

                DictGroupRequired[validateAttrsWrapper.RequiredGroupName].Add(cellStringValue);
            }
        }

        public override ImportInfoWrapper FilterAfterColumnLoop(ImportRowWrapper myRow, IMice_ExcelImportRepository mice_ExcelImportRepository)
        {
            foreach (var item in DictGroupRequired)
            {
                bool isAllFillOrEmpty = item.Value.All(e => string.IsNullOrWhiteSpace(e)) ||
                    item.Value.All(e => !string.IsNullOrWhiteSpace(e));

                if (!isAllFillOrEmpty)
                {
                    ImportInfoWrapper.AddError($"第{myRow.RowNumber}行，{item.Key}不完整", myRow);
                }
            }

            return ImportInfoWrapper;
        }

        public override void OnRowLoopEnd(ImportRowWrapper myRow, IMice_ExcelImportRepository mice_ExcelImportRepository)
        {
            //每行结束后清除数据
            DictGroupRequired.Clear();
        }
    }
}
