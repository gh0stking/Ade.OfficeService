using System;
using System.Collections.Generic;
using System.Linq;

namespace Ade.OfficeService.Excel
{
    public class MobileValidator : BaseValidator
    {
        protected override Type SetAttribute()
        {
            return typeof(MobileAttribute);
        }

        protected override Func<ExcelDataCol, bool> SetDelegateValidate()
        {
            return ValidateMethodCollection.IsMobile;
        }
    }
}
