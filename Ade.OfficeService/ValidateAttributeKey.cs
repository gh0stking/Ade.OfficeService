using System;

namespace Ade.OfficeService.Excel
{
    public class ValidateAttributeKey
    {
        public Type ImportDTOType { get; set; }
        public string ExcelColName { get; set; }

        public override bool Equals(object obj)
        {
            if (obj == null)
            {
                return false;
            }

            ValidateAttributeKey key = (ValidateAttributeKey)obj;
            return key.ImportDTOType == this.ImportDTOType && key.ExcelColName == this.ExcelColName;
        }

        public override int GetHashCode()
        {
            int hash = 13;
            hash = (hash * 7) + ImportDTOType.GetHashCode();
            hash = (hash * 7) + ExcelColName.GetHashCode();

            return hash;
        }
    }
}
