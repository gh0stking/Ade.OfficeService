using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.Excel
{
    //限定字符串校验
    public class RestrictStringAttribute : BaseValidateAttribute
    {
        public RestrictStringAttribute(string restrictStrings)
        {
            this.RestrictStrings = restrictStrings;
        }

        public string RestrictStrings { get; set; }
    }
}
