using System;
using System.Collections.Generic;
using System.Text;
using QDMice.Api.Impl.Repositories;

namespace Ade.OfficeService.Excel
{
    [FilterOption(Order = 7)]
    public class ValidateRestrictString : BaseImportValidator
    {
        public override ImportInfoWrapper FilterInColumnLoop(string cellStringValue, ValidateAttrsWrapper validateAttrsWrapper, ImportRowWrapper myRow, string excelName, IMice_ExcelImportRepository mice_ExcelImportRepository)
        {
            if (validateAttrsWrapper.ValidateRestrictString && !string.IsNullOrWhiteSpace(cellStringValue)
                        && !IsInRestrictStrings(validateAttrsWrapper.RestrictStrings, cellStringValue))
            {
                ImportInfoWrapper.AddError($"第{myRow.RowNumber}行，{excelName}必须为{validateAttrsWrapper.RestrictStrings}", myRow);
            }

            return ImportInfoWrapper;
        }

        private bool IsInRestrictStrings(string restrictStrings, string cellValue)
        {
            string[] arr = restrictStrings.Split(',');
            for (int j = 0; j < arr.Length; j++)
            {
                if (arr[j] == cellValue)
                {
                    return true;
                }
            }

            return false;
        }
    }
}
