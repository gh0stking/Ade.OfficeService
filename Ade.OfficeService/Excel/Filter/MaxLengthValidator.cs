using System;

namespace Ade.OfficeService.Excel
{
    public class MaxLengthValidator : BaseValidator
    {
        protected override Type SetAttribute()
        {
            return typeof(IDNumberAttribute);
        }

        protected override Func<ExcelCol, bool> SetDelegateValidate()
        {
            return ValidateMethodCollection.IsIDNumber;
        }
    }
}
