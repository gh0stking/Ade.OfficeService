using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
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

        public static T FastConvert<T>(ExcelDataRow dataRow)
        {
            string propertyNames = string.Empty;
            dataRow.DataCols.ForEach(c => propertyNames += c.PropertyName + "_");
            var key = typeof(T) + propertyNames.Trim('_');


            if (!Table.ContainsKey(key))
            {
                var dataCols = dataRow.DataCols;

                var dataColsExpr = Expression.Parameter(typeof(List<ExcelDataCol>));

                Expression.Assign(dataColsExpr, Expression.Constant(dataCols));

                List<MemberBinding> memberBindingList = new List<MemberBinding>();

                MethodInfo singleOrDefaultMethod = typeof(Enumerable)
                                                            .GetMethods()
                                                            .Single(m => m.Name == "SingleOrDefault" && m.GetParameters().Count() == 2)
                                                            .MakeGenericMethod(new[] { typeof(ExcelDataCol) });

                foreach (var prop in typeof(T).GetProperties())
                {
                    //ChangeType(dataRow.DataCols.SingleOrDefault(d => d.PropertyName == prop.Name)?.ColValue, prop.PropertyType);

                    Expression<Func<ExcelDataCol, bool>> lambdaExpr = c => c.PropertyName == prop.Name;

                    MethodCallExpression singleOrDefaultExpr = Expression.Call(
                        singleOrDefaultMethod
                        , dataColsExpr
                        , lambdaExpr);

                    var colValueExpr = Expression.Parameter(typeof(string));
                    Expression.Assign(colValueExpr, Expression.Property(singleOrDefaultExpr, "ColValue"));

                    MethodInfo changeTypeMethod = typeof(ExpressionMapper).GetMethods().Where(m => m.Name == "ChangeType").First();

                    var objColValueExpr = Expression.Parameter(typeof(object));

                    Expression expr =
                        Expression.Convert(
                            Expression.Call(changeTypeMethod
                                , Expression.Property(
                                    Expression.Call(
                                          singleOrDefaultMethod
                                        , Expression.Constant(dataRow.DataCols)
                                        , lambdaExpr)
                                        , typeof(ExcelDataCol), "ColValue"), Expression.Constant(prop.PropertyType))
                                    , prop.PropertyType);

                    memberBindingList.Add(Expression.Bind(prop, expr));
                }

                MemberInitExpression memberInitExpression = Expression.MemberInit(Expression.New(typeof(T)), memberBindingList.ToArray());
                Expression<Func<ExcelDataRow, T>> lambda = Expression.Lambda<Func<ExcelDataRow, T>>(memberInitExpression, new ParameterExpression[]
                {
                    Expression.Parameter(typeof(ExcelDataRow), "p")
                });

                Func<ExcelDataRow, T> func = lambda.Compile();//拼装是一次性的
                Table[key] = func;
            }
            var ss = (Func<ExcelDataRow, T>)Table[key];

            return ((Func<ExcelDataRow, T>)Table[key]).Invoke(dataRow);
        }

        //private static Expression<Func<string, T>> CreateLambda<T>()
        //{
        //    var stringValue = Expression.Parameter(
        //        typeof(T), "stringValue");

        //    var call = Expression.Call(
        //        typeof(Enumerable), "GenericChangeType", new Type[] { typeof(T) }, stringValue);

        //    return Expression.Lambda<Func<string, T>>(call, stringValue);
        //}

        /// <summary>
        /// 字典缓存表达式树
        /// </summary>
        /// <typeparam name="TIn"></typeparam>
        /// <typeparam name="TOut"></typeparam>
        /// <param name="dataRow"></param>
        /// <returns></returns>
        internal static TOut Trans<TOut>(ExcelDataRow dataRow)
        {
            string key = string.Format("funckey_{0}_{1}", typeof(ExcelDataRow).FullName, typeof(TOut).FullName);
            //if (!Table.ContainsKey(key))
            {
                ParameterExpression parameterExpression = Expression.Parameter(typeof(ExcelDataRow), "p");
                List<MemberBinding> memberBindingList = new List<MemberBinding>();
                foreach (var item in typeof(TOut).GetProperties())
                {
                    string stringValue = dataRow.DataCols.SingleOrDefault(c => c.PropertyName == item.Name)?.ColValue;
                    ConstantExpression constant = Expression.Constant(ChangeType(stringValue, item.PropertyType));
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
                Expression<Func<ExcelDataRow, TOut>> lambda = Expression.Lambda<Func<ExcelDataRow, TOut>>(memberInitExpression, new ParameterExpression[]
                {
                    parameterExpression
                });
                Func<ExcelDataRow, TOut> func = lambda.Compile();//拼装是一次性的
                Table[key] = func;
            }
            return ((Func<ExcelDataRow, TOut>)Table[key]).Invoke(dataRow);
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
