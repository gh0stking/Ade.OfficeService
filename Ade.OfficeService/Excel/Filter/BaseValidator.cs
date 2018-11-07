using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Ade.OfficeService.Excel
{
    public abstract class BaseValidator : IFilter
    {
        /// <summary>
        /// 与校验器关联的特性类型
        /// </summary>
        protected Type AttributeType { get => this.attributeType; }
        private Type attributeType;

        /// <summary>
        /// 需要校验的Excel列集合
        /// </summary>
        //protected List<ExcelCol> MarkedCols { get => this.markedCols; }
        //private List<ExcelCol> markedCols;

        /// <summary>
        /// Excel数据集合
        /// </summary>
        protected List<ExcelDataRow> DataRows { get => this.dataRows; }
        private List<ExcelDataRow> dataRows;

        /// <summary>
        /// 校验失败提示语
        /// </summary>
        protected string ErrorMsg { get => this.errorMsg; }
        private string errorMsg;

        /// <summary>
        /// 校验委托
        /// </summary>
        protected Func<ExcelDataCol, bool> DelegateValidate { get => this.delegateValidate; }
        private Func<ExcelDataCol, bool> delegateValidate;

        /// <summary>
        /// 数据库校验委托(参数为表名，字段名)
        /// </summary>
        protected Func<string, string, bool> DelegateDatabaseValidate { get => this.delegateDatabaseValidate; }
        private Func<string, string, bool> delegateDatabaseValidate;

        /// <summary>
        /// DTO属性的特性标记字典
        /// </summary>
        protected Dictionary<string,BaseValidateAttribute> ValidateAttributes { get => this.validateAttributes; }
        private Dictionary<string, BaseValidateAttribute> validateAttributes;


        protected Type ImportDTOType { get => this.importDTOType; }
        private Type importDTOType;


        /// <summary>
        /// 注册校验器
        /// </summary>
        /// <param name="importDTOType">导入DTO类型</param>
        /// <param name="headerRow">表头</param>
        /// <param name="dataRows">数据列集合</param>
        /// <param name="errorMsg">校验失败提示语</param>
        public void Register(Type importDTOType, ExcelHeaderRow headerRow, List<ExcelDataRow> dataRows, string errorMsg, Func<string, string, bool> delegateDatabaseValidate = null)
        {
            if (importDTOType == null)
            {
                throw new ArgumentNullException("importDTOType");
            }

            if (attributeType == null)
            {
                throw new ArgumentNullException("attributeType");
            }

            if (headerRow == null)
            {
                throw new ArgumentNullException("headerRow");
            }

            this.dataRows = dataRows;
            this.errorMsg = errorMsg;
            this.attributeType = SetAttribute();
            this.importDTOType = importDTOType;
            //this.markedCols = SetMarkedExcelCols(importDTOType, this.attributeType, headerRow);
            this.delegateValidate = SetDelegateValidate();
            this.delegateDatabaseValidate = delegateDatabaseValidate;
        }

        /// <summary>
        /// 设置与校验器关联的特性类型
        /// </summary>
        /// <returns></returns>
        protected abstract Type SetAttribute();

        /// <summary>
        /// 设置需要校验的Excel列
        /// </summary>
        /// <returns></returns>
        //private List<ExcelCol> SetMarkedExcelCols(Type importDTOType, Type attributeType, ExcelHeaderRow headerRow)
        //{
        //    if (importDTOType == null)
        //    {
        //        throw new ArgumentNullException("importDTOType");
        //    }

        //    if (attributeType == null)
        //    {
        //        throw new ArgumentNullException("attributeType");
        //    }

        //    if (headerRow == null)
        //    {
        //        throw new ArgumentNullException("headerRow");
        //    }

        //    List<ExcelCol> excelCols = new List<ExcelCol>();
           
        //    importDTOType.GetProperties().ToList().ForEach((p) => {
        //        if (p.IsDefined(attributeType.GetType()))
        //        {
        //            excelCols.AddRange(headerRow.HeaderRow.Where(e => e.ColName.Trim().Equals(p.Name.Trim(), StringComparison.OrdinalIgnoreCase)));
        //        }
        //    });

        //    return excelCols;
        //}

        /// <summary>
        /// 设置校验方法
        /// </summary>
        /// <returns></returns>
        protected virtual Func<ExcelDataCol, bool> SetDelegateValidate()
        {
            return null;
        }

        protected virtual Func<string, string, bool> SetDatabaseDelegateValidate()
        {
            return null;
        }

        /// <summary>
        /// 校验
        /// </summary>
        /// <returns></returns>
        public virtual List<ExcelDataRow> Filter
        {
            if (this.DelegateValidate == null)
            {
                throw new ArgumentNullException("DelegateValidate");
            }

            this.DataRows.Where(r => r.IsValid).ToList().ForEach((r) =>
            {
                r.DataCols.ForEach(c => {
                    var attr = ValidateAttributeFactory.Create(this.AttributeType, this.ImportDTOType, c.ColName);
                    if (attr != null)
                    {
                        r.SetState(this.DelegateValidate(c), c, this.ErrorMsg);
                    }
                });
            });

            return this.DataRows;
        }
    }
}
