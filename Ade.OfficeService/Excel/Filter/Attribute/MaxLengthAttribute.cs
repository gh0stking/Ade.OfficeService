using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.Excel
{
    //最大长度
    public class MaxLengthAttribute : BaseValidateAttribute
    {
        public MaxLengthAttribute(int maxLength)
        {
            this.MaxLength = maxLength;
        }

        public int MaxLength { get; set; }
    }
}
