using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.Excel
{
    //分组必填校验
    public class GroupRequiredAttribute : BaseValidateAttribute
    {
        public GroupRequiredAttribute(string groupName)
        {
            this.RequiredGroupName = groupName;
        }

        public string RequiredGroupName { get; set; }
    }
}
