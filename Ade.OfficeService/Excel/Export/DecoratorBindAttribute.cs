using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.Excel
{
    public class DecoratorBindAttribute : Attribute
    {
        public DecoratorBindAttribute(Type decoratorAttributeType)
        {
            this.DecoratorAttributeType = decoratorAttributeType;
        }

        public Type DecoratorAttributeType { get; set; }
    }
}
