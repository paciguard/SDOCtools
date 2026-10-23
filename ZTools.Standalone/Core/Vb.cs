using System;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ZTools.Engine
{
    // Preserve VBA's Empty, string concatenation, mixed numeric/string comparisons,
    // one-based Mid, and UTF-16 LenB while moving execution entirely into C#.
    internal static class Vb
    {
        public static string Text(object value) => value == null ? "" : Conversions.ToString(value);
        public static double Num(object value) => value == null ? 0 : Conversions.ToDouble(value);
        public static int Int(object value) => Conversions.ToInteger(value);
        public static bool Bool(object value) => value != null && Conversions.ToBoolean(value);
        public static dynamic Add(object a, object b) => Operators.AddObject(a, b);
        public static dynamic Sub(object a, object b) => Operators.SubtractObject(a, b);
        public static dynamic Mul(object a, object b) => Operators.MultiplyObject(a, b);
        public static dynamic Div(object a, object b) => Operators.DivideObject(a, b);
        public static dynamic Pow(object a, object b) => Operators.ExponentObject(a, b);
        public static dynamic Neg(object a) => Operators.NegateObject(a);
        public static string Cat(object a, object b) => Text(a) + Text(b);
        public static bool Eq(object a, object b) => Operators.ConditionalCompareObjectEqual(a, b, false);
        public static bool Ne(object a, object b) => !Eq(a, b);
        public static bool Lt(object a, object b) => Operators.ConditionalCompareObjectLess(a, b, false);
        public static bool Gt(object a, object b) => Operators.ConditionalCompareObjectGreater(a, b, false);
        public static bool Le(object a, object b) => Operators.ConditionalCompareObjectLessEqual(a, b, false);
        public static bool Ge(object a, object b) => Operators.ConditionalCompareObjectGreaterEqual(a, b, false);
        public static dynamic And(object a, object b) => Operators.AndObject(a, b);
        public static dynamic Or(object a, object b) => Operators.OrObject(a, b);
        public static bool Not(object a) => !Bool(a);
        public static int Len(object s) => Text(s).Length;
        public static int Lenb(object s) => Text(s).Length * 2;
        public static string Left(object s, object n) => Strings.Left(Text(s), Int(n));
        public static string Right(object s, object n) => Strings.Right(Text(s), Int(n));
        public static string Mid(object s, object start, object count) => Strings.Mid(Text(s), Int(start), Int(count));
        public static int Instr(object s, object part) => Strings.InStr(Text(s), Text(part), CompareMethod.Binary);
        public static double Val(object s) => Conversion.Val(Text(s));
        public static int Asc(object s) => Strings.Asc(Text(s));
        public static string Replace(object s, object find, object replacement) => Strings.Replace(Text(s), Text(find), Text(replacement));
        public static string Ucase(object s) => Text(s).ToUpperInvariant();
        public static string Trim(object s) => Text(s).Trim();
        public static string Cstr(object s) => Text(s);
        public static int Cint(object s) => Int(s);
        public static bool Isnumeric(object s) => Information.IsNumeric(s);
        public static string Typename(object s) => Information.TypeName(s);
        public static string Format(object s, object f) => Strings.Format(s, Text(f));
        public static double[] Resize(object array, int count)
        {
            var result = new double[count];
            Array.Copy((Array)array, result, Math.Min(((Array)array).Length, count));
            return result;
        }
    }
}
