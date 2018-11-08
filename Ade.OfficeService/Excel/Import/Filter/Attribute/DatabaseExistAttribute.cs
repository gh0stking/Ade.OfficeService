﻿using System;

namespace Ade.OfficeService.Excel
{
    //数据库已存在数据校验
    [AttributeUsage(AttributeTargets.Property,AllowMultiple =true,Inherited =false)]
    public class DatabaseExistAttribute : BaseFilterAttribute
    {
        public DatabaseExistAttribute(string tableName, string fieldName)
        {
            this.TableName = tableName;
            this.FieldName = fieldName;
            this.ErrorMsg = "数据已存在";
        }

        public string TableName { get; set; }
        public string FieldName { get; set; }
    }
}
