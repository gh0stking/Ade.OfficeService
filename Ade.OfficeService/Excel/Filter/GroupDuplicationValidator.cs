using System;
using System.Collections.Generic;
using System.Linq;

namespace Ade.OfficeService.Excel
{
    public class ValidateGroupDuplication : BaseValidator
    {
        protected override Type SetAttribute()
        {
            return typeof(GroupDuplicationAttribute);
        }

        public override List<ExcelDataRow> Validate()
        {
            this.DataRows.Where(r => r.IsValid).ToList().ForEach((r) =>
            {
                r.DataCols.ForEach(c => {
                    var attr = ValidateAttrFactory.Create(this.AttributeType, this.ImportDTOType, c.ColName);
                    if (attr != null)
                    {

                    }
                });
            });

            return this.DataRows;

        }
    }
}
