namespace Ade.OfficeService.Excel
{
    //数据库已存在数据校验
    public class DBExistAttribute : BaseValidateAttribute
    {
        public DBExistAttribute(string tableName, string fieldName)
        {
            this.TableName = tableName;
            this.FieldName = fieldName;
        }

        public string TableName { get; set; }
        public string FieldName { get; set; }
    }
}
