using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConEd.HeavyBid.Importer.Utility
{
    public class ComReleaser : IDisposable
    {
        Stack<object> objects = new Stack<object>();
        public T Add<T>(Expression<Func<T>> func)
        {
            return (T)Walk(func.Body);
        }
        object Walk(Expression expr)
        {
            object obj = WalkImpl(expr);
            if (obj != null && Marshal.IsComObject(obj) && !objects.Contains(obj))
            {
                objects.Push(obj);
            }
            return obj;
        }
        object[] Walk(IEnumerable<Expression> args)
        {
            if (args == null) return null;
            return args.Select(arg => Walk(arg)).ToArray();
        }
        object WalkImpl(Expression expr)
        {
            switch (expr.NodeType)
            {
                case ExpressionType.Constant:
                    return ((ConstantExpression)expr).Value;
                case ExpressionType.New:
                    NewExpression ne = (NewExpression)expr;
                    return ne.Constructor.Invoke(Walk(ne.Arguments));
                case ExpressionType.MemberAccess:
                    MemberExpression me = (MemberExpression)expr;
                    object target = Walk(me.Expression);
                    switch (me.Member.MemberType)
                    {
                        case MemberTypes.Field:
                            return ((FieldInfo)me.Member).GetValue(target);
                        case MemberTypes.Property:
                            return ((PropertyInfo)me.Member).GetValue(target, null);
                        default:
                            throw new NotSupportedException();

                    }
                case ExpressionType.Call:
                    MethodCallExpression mce = (MethodCallExpression)expr;
                    return mce.Method.Invoke(Walk(mce.Object), Walk(mce.Arguments));
                default:
                    throw new NotSupportedException();
            }
        }
        public void Dispose()
        {
            while (objects.TryPop(out var obj))
            {
                if (obj is Excel.Application app)
                {
                    app.Quit();
                }
                if (obj is Excel.Workbooks wbs)
                {
                    wbs.Close();
                }
                if (obj is Excel.Workbook wb)
                {
                    wb.Close();
                }
                Marshal.ReleaseComObject(obj);
                Debug.WriteLine("Released: " + obj);
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
