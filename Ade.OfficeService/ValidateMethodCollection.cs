using NPOI.SS.UserModel;
using System;
using System.Text.RegularExpressions;

namespace Ade.OfficeService.Excel
{
    public static class ValidateMethodCollection
    {
        /// <summary>
        /// 是否正确手机号
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static bool IsMobile(ExcelDataCol cell)
        {
            return Regex.IsMatch(cell.ColValue, @"1\d{10}$");
        }

        /// <summary>
        /// 是否正确车牌号
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static bool IsCarCode(ExcelDataCol cell)
        {
            return Regex.IsMatch(cell.ColValue, @"^[京津沪渝冀豫云辽黑湘皖鲁新苏浙赣鄂桂甘晋蒙陕吉闽贵粤青藏川宁琼使领A-Z]{1}[A-Z]{1}[A-Z0-9]{4}[A-Z0-9挂学警港澳]{1}$");
        }

        /// <summary>
        /// 是否身份证号
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static bool IsIDNumber(ExcelDataCol cell)
        {
            return Regex.IsMatch(cell.ColValue, @"^(^\d{15}$|^\d{18}$|^\d{17}(\d|X|x))$", RegexOptions.IgnoreCase);
        }

        /// <summary>
        /// 是否为日期
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static bool IsDateTime(ExcelDataCol cell)
        {
            bool isDateTime = true;
            try
            {
                DateTime.Parse(cell.ColValue);
            }
            catch (Exception)
            {
                isDateTime = false;
            }

            return isDateTime;
        }

        /// <summary>
        /// 是否为空
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static bool IsHasValue(ExcelDataCol cell)
        {
            return !string.IsNullOrWhiteSpace(cell.ColValue);
        }

        /// <summary>
        /// 是否没有超出字符串长度范围
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static bool IsNotBeyondMaxLength(ExcelDataCol cell)
        {
            return string.IsNullOrWhiteSpace(cell.ColValue);
        }
    }
}
