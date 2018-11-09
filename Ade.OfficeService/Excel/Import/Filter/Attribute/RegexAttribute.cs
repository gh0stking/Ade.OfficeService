using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace Ade.OfficeService.Excel
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true, Inherited = false)]
    public class RegexAttribute : BaseFilterAttribute
    {
        public string Regex { get; set; }
        public RegexAttribute(string regex)
        {
            this.Regex = regex;
            this.ErrorMsg = "非法";
        }

        public RegexAttribute(RegexEnum regexEnum)
        {
            this.Regex = RegexFactory.CreateRegex(regexEnum);
        }
    }
}
