using System;

namespace Ade.OfficeService.Excel
{
    public class CarCodeValidator : BaseValidator
    {
        protected override Type SetAttribute()
        {
            return typeof(CarCodeAttribute);
        }

        protected override Func<ExcelDataCol, bool> SetDelegateValidate()
        {
            return ValidateMethodCollection.IsCarCode;
        }
    }
}
