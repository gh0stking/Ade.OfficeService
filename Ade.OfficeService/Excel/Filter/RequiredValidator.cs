using System;
using System.Collections.Generic;
using System.Linq;

namespace Ade.OfficeService.Excel
{
    [FilterOption(Priority =PriorityEnum.Critical)]
    public class RequiredFilter : BaseValidator
    {
        protected override Type SetAttribute()
        {
            return typeof(RequiredAttribute);
        }

        protected override Func<ExcelDataCol, bool> SetDelegateValidate()
        {
            return ValidateMethodCollection.IsHasValue;
        }
    }
}
