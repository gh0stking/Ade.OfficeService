using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.Excel
{
    //分组重复数据校验
    //多字段校验是否重复，如手机号和员工类型都相同才判定重复
    public class GroupDuplicationAttribute : BaseValidateAttribute
    {
        public GroupDuplicationAttribute()
        {
            this.DuplicationGroupName = "分组重复校验";
        }

        public GroupDuplicationAttribute(string duplicationGroupName)
        {
            this.DuplicationGroupName = duplicationGroupName;
        }

        public string DuplicationGroupName { get; set; }
    }
}
