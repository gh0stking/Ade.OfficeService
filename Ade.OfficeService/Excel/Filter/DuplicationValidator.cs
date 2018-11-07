using System;
using System.Collections.Generic;
using System.Linq;

namespace Ade.OfficeService.Excel
{
    public class ValidateDuplication : BaseValidator
    {
        protected override Type SetAttribute()
        {
            return typeof(DuplicationAttribute);
        }

        public override List<ExcelDataRow> Validate()
        {
            Dictionary<int, string> dictColIndexAndValue = new Dictionary<int, string>();
            this.DataRows.Where(r => r.IsValid).ToList().ForEach((r) =>
            {
                r.DataCols.ForEach(c => {
                    var attr = ValidateAttrFactory.Create(this.AttributeType, this.ImportDTOType, c.ColName);
                    if (attr != null)
                    {
                        r.SetState(dictColIndexAndValue[c.ColIndex].Contains(c.ColValue), c, this.ErrorMsg);
                        dictColIndexAndValue.Add(c.ColIndex, c.ColValue);
                    }
                });
            });

            return this.DataRows;
        }
    }
}
