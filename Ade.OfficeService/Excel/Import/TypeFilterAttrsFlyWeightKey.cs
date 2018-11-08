using System;

namespace Ade.OfficeService.Excel
{
    public class TypeFilterAttrsFlyWeightKey
    {
        public Type ImportDTOType { get; set; }
        public ExcelHeaderRow ExcelHeaderRow { get; set; }

        public override bool Equals(object obj)
        {
            if (obj == null)
            {
                return false;
            }

            TypeFilterAttrsFlyWeightKey key = (TypeFilterAttrsFlyWeightKey)obj;
            return key.ImportDTOType == this.ImportDTOType && key.ExcelHeaderRow == this.ExcelHeaderRow;
        }

        public override int GetHashCode()
        {
            int hash = 13;
            hash = (hash * 7) + ImportDTOType.GetHashCode();
            hash = (hash * 7) + ExcelHeaderRow.GetHashCode();

            return hash;
        }
    }
}
