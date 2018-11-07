using QDMice.Api.Impl.Repositories;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ade.OfficeService.Excel
{
    [FilterOption(Priority = PriorityEnum.Low)]
    public class ValidateDBExist : BaseValidator
    {
        protected override Type SetAttribute()
        {
            return typeof(DBExistAttribute);
        }

        public override List<ExcelDataRow> Validate()
        {
            if (this.DelegateDatabaseValidate == null)
            {
                throw new ArgumentNullException("DelegateDatabaseValidate");
            }

            this.DataRows.Where(r => r.IsValid).ToList().ForEach((r) =>
            {
                r.DataCols.ForEach(c => {
                    var attr = (DBExistAttribute)ValidateAttrFactory.Create(this.AttributeType, this.ImportDTOType, c.ColName);
                    if(attr != null)
                    {
                        r.SetState(this.DelegateDatabaseValidate(attr.TableName, attr.FieldName), c, this.ErrorMsg);
                    }
                });
            });

            return this.DataRows;

        }
    }
}
