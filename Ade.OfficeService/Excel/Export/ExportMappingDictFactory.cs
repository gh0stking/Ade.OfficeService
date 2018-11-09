﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Ade.OfficeService.Excel
{
    public static class ExportMappingDictFactory
    {
        private static readonly Hashtable Table = Hashtable.Synchronized(new Hashtable(1024));

        /// <summary>
        /// 创建映射字段，键：DTO属性名，值：Excel列名
        /// </summary>
        /// <param name="exportType"></param>
        /// <returns></returns>
        public static Dictionary<string, string> CreateInstance(Type exportType)
        {
            var key = exportType;
            if (Table[key] != null)
            {
                return (Dictionary<string, string>)Table[key];
            }

            Dictionary<string, string> dict = new Dictionary<string, string>();
            exportType.GetProperties().ToList().ForEach(p=> {
                if (p.IsDefined(typeof(ExcelExportAttribute)))
                {
                    dict.Add(p.Name, p.GetCustomAttribute<ExcelExportAttribute>().ColName);
                }
            });

            Table[key] = dict;
            return dict;
        }
            
    }
}
