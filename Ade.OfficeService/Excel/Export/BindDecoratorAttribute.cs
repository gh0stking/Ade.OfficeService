using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.Excel
{
    public class BindDecoratorAttribute : Attribute
    {
        public BindDecoratorAttribute(Type decoratorType)
        {
            this.DecoratorType = decoratorType;
        }

        public Type DecoratorType { get; set; }
    }
}
