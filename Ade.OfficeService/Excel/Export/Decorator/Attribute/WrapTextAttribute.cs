using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.Excel
{
    /// <summary>
    /// 自动换行
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
    public class WrapTextAttribute : BaseDecorateAttribute
    {
    }
}
