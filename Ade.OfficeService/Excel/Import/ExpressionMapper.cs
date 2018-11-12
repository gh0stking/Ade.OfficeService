using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace Ade.OfficeService.Excel
{
    /// <summary>
    /// 生成表达式目录树 缓存
    /// </summary>
    public class ExpressionMapper
    {
        private static Hashtable Table = Hashtable.Synchronized(new Hashtable(1024));

        /// <summary>
        /// 字典缓存表达式树
        /// </summary>
        /// <typeparam name="TIn"></typeparam>
        /// <typeparam name="TOut"></typeparam>
        /// <param name="dataRow"></param>
        /// <returns></returns>
        public static TOut Trans<TOut>(ExcelDataRow dataRow)
        {
            string key = string.Format("funckey_{0}_{1}", typeof(ExcelDataRow).FullName, typeof(TOut).FullName);
            //if (!Table.ContainsKey(key))
            {
                ParameterExpression parameterExpression = Expression.Parameter(typeof(ExcelDataRow), "p");
                List<MemberBinding> memberBindingList = new List<MemberBinding>();
                foreach (var item in typeof(TOut).GetProperties())
                {
                    string stringValue = dataRow.DataCols.SingleOrDefault(c => c.PropertyName == item.Name)?.ColValue;
                    ConstantExpression constant = Expression.Constant(ChangeType(stringValue,item.PropertyType));
                    MemberBinding memberBinding = Expression.Bind(item, constant);
                    memberBindingList.Add(memberBinding);
                }
                foreach (var item in typeof(TOut).GetFields())
                {
                    string stringValue = dataRow.DataCols.SingleOrDefault(c => c.PropertyName == item.Name)?.ColValue;
                    ConstantExpression constant = Expression.Constant(ChangeType(stringValue, item.FieldType));
                    MemberBinding memberBinding = Expression.Bind(item, constant);
                    memberBindingList.Add(memberBinding);
                }
                MemberInitExpression memberInitExpression = Expression.MemberInit(Expression.New(typeof(TOut)), memberBindingList.ToArray());
                Expression<Func<ExcelDataRow,TOut>> lambda = Expression.Lambda<Func<ExcelDataRow,TOut>>(memberInitExpression, new ParameterExpression[]
                {
                    parameterExpression
                });
                Func<ExcelDataRow,TOut> func = lambda.Compile();//拼装是一次性的
                Table[key] = func;
            }
            return ((Func<ExcelDataRow,TOut>)Table[key]).Invoke(dataRow);
        }

        public static object ChangeType(string stringValue, Type type)
        {
            object obj = null;

            Type nullableType = Nullable.GetUnderlyingType(type);
            if (nullableType != null)
            {
                if (stringValue == null)
                {
                    obj = null;
                }

            }
            else if (typeof(System.Enum).IsAssignableFrom(type))
            {
                obj = Enum.Parse(type, stringValue);
            }
            else
            {
                obj = Convert.ChangeType(stringValue, type);
            }

            return obj;
        }
    }
}
