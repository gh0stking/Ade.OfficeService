using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.Excel
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class MergeColsAttribute : BaseDecorateAttribute
    {
    }
}
