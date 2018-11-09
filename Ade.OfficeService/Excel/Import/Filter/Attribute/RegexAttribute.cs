using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace Ade.OfficeService.Excel
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true, Inherited = false)]
    public class RegexAttribute : BaseFilterAttribute
    {
        public string Regex { get; set; }
        public bool IsPreDefined { get; set; } = false;

        public RegexAttribute(string regex)
        {
            this.Regex = regex;
            this.ErrorMsg = "非法";
        }

        public RegexAttribute(bool isPredefined, PreDefinedRegexEnum regexEnum)
        {
            if (isPredefined)
            {
                this.Regex = PredefineRegexFactory.GetPredefineRegex(regexEnum);
            }
        }
    }

    public enum PreDefinedRegexEnum
    {
        车牌号 = 10,
        身份证号 = 20,
        性别 = 30,
        国内手机号 = 40,
        邮箱 = 50,
        非空 = 60,
        网址URL = 70,
    }

    public static class PredefineRegexFactory
    {
        public static string GetPredefineRegex(PreDefinedRegexEnum regexEnum)
        {
            string regex = string.Empty;
            switch (regexEnum)
            {
                case PreDefinedRegexEnum.车牌号:
                    regex = RegexConstant.CAR_CODE_REGEX;
                    break;
                case PreDefinedRegexEnum.身份证号:
                    regex = RegexConstant.IDENTITY_NUMBER_REGEX;
                    break;
                case PreDefinedRegexEnum.性别:
                    regex = RegexConstant.GENDER_REGEX;
                    break;
                case PreDefinedRegexEnum.国内手机号:
                    regex = RegexConstant.MOBILE_CHINA_REGEX;
                    break;
                case PreDefinedRegexEnum.邮箱:
                    regex = RegexConstant.EMAIL_REGEX;
                    break;
                case PreDefinedRegexEnum.非空:
                    regex = RegexConstant.NOT_EMPTY_REGEX;
                    break;
                case PreDefinedRegexEnum.网址URL:
                    regex = RegexConstant.URL_REGEX;
                    break;
                default:
                    break;
            }

            return regex;
        }
    }
}
