using System;

namespace Ade.OfficeService.Excel
{
    public class DateTimeValidator : BaseValidator
    {
        protected override Type SetAttribute()
        {
            return typeof(DateTimeAttribute);
        }

        protected override Func<ExcelDataCol, bool> SetDelegateValidate()
        {
            return ValidateMethodCollection.IsDateTime;
        }
    }
}
