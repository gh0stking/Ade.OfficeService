using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.Excel
{ 
    public class FilterOptionAttribute : BaseValidateAttribute
    {
        public PriorityEnum Priority { get; set; } = PriorityEnum.Mediumn;
        public bool Disable { get; set; } = false;
    }
}
