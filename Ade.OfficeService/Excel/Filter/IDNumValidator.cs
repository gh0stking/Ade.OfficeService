using System;

namespace Ade.OfficeService.Excel
{
    public class IDNumValidator : BaseValidator
    {
        protected override Type SetAttribute()
        {
            return typeof(IDNumberAttribute);
        }

        protected override Func<ExcelDataCol, bool> SetDelegateValidate()
        {
            return ValidateMethodCollection.IsIDNumber;
        }
    }
}
