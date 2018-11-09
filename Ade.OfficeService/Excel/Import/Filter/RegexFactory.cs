using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.Excel
{
    public static class RegexFactory
    {
        public static string CreateRegex(RegexEnum regexEnum)
        {
            string regex = string.Empty;
            switch (regexEnum)
            {
                case RegexEnum.车牌号:
                    regex = RegexConstant.CAR_CODE_REGEX;
                    break;
                case RegexEnum.身份证号:
                    regex = RegexConstant.IDENTITY_NUMBER_REGEX;
                    break;
                case RegexEnum.性别:
                    regex = RegexConstant.GENDER_REGEX;
                    break;
                case RegexEnum.国内手机号:
                    regex = RegexConstant.MOBILE_CHINA_REGEX;
                    break;
                case RegexEnum.邮箱:
                    regex = RegexConstant.EMAIL_REGEX;
                    break;
                case RegexEnum.非空:
                    regex = RegexConstant.NOT_EMPTY_REGEX;
                    break;
                case RegexEnum.网址URL:
                    regex = RegexConstant.URL_REGEX;
                    break;
                default:
                    break;
            }

            return regex;
        }
    }
}
