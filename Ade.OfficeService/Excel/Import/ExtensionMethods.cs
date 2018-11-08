using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Ade.OfficeService.Excel
{
    public static class ExtensionMethods
    {
        /// <summary>
        /// 获取某单元格的某校验特性
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="col"></param>
        /// <param name="typeAttrs"></param>
        /// <returns></returns>
        public static T GetFilterAttr<T>(this ExcelDataCol col, TypeFilterAttrsFlyWeight typeAttrs)
            where T : BaseFilterAttribute
        {
            return (T)typeAttrs.PropertyFilterAttrs.SingleOrDefault(a => a.ColIndex == col.ColIndex)?.
                 FilterAttrs?.SingleOrDefault(e => e.GetType() == typeof(T));
        }

        /// <summary>
        /// 设置Excel行的校验结果
        /// </summary>
        /// <param name="row"></param>
        /// <param name="isValid"></param>
        /// <param name="col"></param>
        /// <param name="errorMsg"></param>
        public static void SetState(this ExcelDataRow row, bool isValid, ExcelDataCol dataCol, string errorMsg)
        {
            if (!isValid)
            {
                row.IsValid = isValid;
                row.ErrorMsgs.Add(dataCol, errorMsg);
            }
        }

        /// <summary>
        /// 是否是日期
        /// </summary>
        /// <param name="col"></param>
        /// <returns></returns>
        public static bool IsDateTime(this ExcelDataCol col)
        {
            return DateTime.TryParse(col.ColValue, out DateTime dt);
        }

        /// <summary>
        /// 判断数值是否在范围内
        /// </summary>
        /// <param name="col"></param>
        /// <param name="max"></param>
        /// <param name="min"></param>
        /// <returns></returns>
        public static bool IsInRange(this ExcelDataCol col, decimal max, decimal min)
        {
            if (!decimal.TryParse(col.ColValue, out decimal val))
            {
                return false;
            }

            if (val > max || val < min)
            {
                return false;
            }

            return true;
        }

        public static string GetStringValue(this ICell cell)
        {
            try
            {
                switch (cell.CellType)
                {
                    case CellType.Boolean:
                        return cell.BooleanCellValue.ToString();
                    case CellType.Error:
                        return cell.ErrorCellValue.ToString();
                    case CellType.Numeric:
                        return DateUtil.IsCellDateFormatted(cell)
                       ? cell.DateCellValue.ToString()
                       : cell.NumericCellValue.ToString();
                    case CellType.String:
                        return cell.StringCellValue;
                    default:
                        return cell.ToString();
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// 将ExcelDataRow转换为指定类型
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="row"></param>
        /// <returns></returns>
        public static T Convert<T>(this ExcelDataRow row)
        {
            Type t = typeof(T);
            object o = Activator.CreateInstance(t);
            t.GetProperties().ToList().ForEach(p => {
                if (p.IsDefined(typeof(ExcelImportAttribute)))
                {
                    ExcelDataCol col = row.DataCols.SingleOrDefault(c => c.ColName == p.GetCustomAttribute<ExcelImportAttribute>().ColName);

                    if (col != null)
                    {
                        p.SetValue(o, GetValue(col, p));
                    }
                }
            });

            return (T)o;
        }

        private static object GetValue(ExcelDataCol col, PropertyInfo prop)
        {
            string cellStringValue = col.ColValue;

            object objValue;

            //是否枚举类型
            if (prop.PropertyType.IsSubclassOf(typeof(Enum)))
            {
                return Enum.Parse(prop.PropertyType, string.IsNullOrWhiteSpace(cellStringValue) ? "0" : cellStringValue);
            }

            switch (prop.PropertyType.Name)
            {
                case "DateTime":
                    objValue = string.IsNullOrWhiteSpace(cellStringValue) ? DateTime.MinValue : DateTime.Parse(cellStringValue);
                    break;
                case "Int32":
                    objValue = string.IsNullOrWhiteSpace(cellStringValue) ? 0 : int.Parse(cellStringValue);
                    break;
                case "Int64":
                    objValue = string.IsNullOrWhiteSpace(cellStringValue) ? 0 : long.Parse(cellStringValue);
                    break;
                case "Boolean":
                    string[] strs = new string[] { "true", "是" };
                    objValue = strs.ToList().Contains(cellStringValue);
                    break;
                case "String":
                    objValue = cellStringValue;
                    break;
                default:
                    throw new NotSupportedException("不支持的数据类型");
            }

            return objValue;
        }
    }
}
