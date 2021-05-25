using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.Registration;

namespace RELATIONAL_ALGEBRA
{
    public static class Rel_Algebra
    {
        private static object[,] PREPROCCESS(object[,] Range, ref bool Headers)
        {
            return REMOVE_DUBLICATES(REMOVE_EMPTY(Range, ref Headers), ref Headers);
        }

        private static object[,] POSTPROCCESS(object[,] Range, ref bool Headers)
        {
            return REMOVE_DUBLICATES(Range, ref Headers);
        }

        [ExcelFunction(Description = "Returns a relational union of two sets as an 2 array.")]
        public static object[,] REL_UNION([ExcelArgument("Required. First range for uniting.")] object[,] Range1,
            [ExcelArgument("Required. Second range for uniting.")] object[,] Range2,
            [ExcelArgument("Optional. Enter TRUE if your ranges have titles, otherwise - FALSE. Default value is TRUE.")] object Headers_arg)
        {
            bool Headers = Optional.Check(Headers_arg, true);
            bool Header1 = Headers;
            bool Header2 = Headers;
            Range1 = PREPROCCESS(Range1, ref Header1);
            Range2 = PREPROCCESS(Range2, ref Header2);
            if (Header1 ^ Header2 == true)
            {
                if (Header1) Range1 = TRIM_ARRAY(0, Range1);
                if (Header2) Range2 = TRIM_ARRAY(0, Range2);
            }
            Headers = Header1 && Header2;

            int rows1 = Range1.GetLength(0);
            int cols1 = Range1.GetLength(1);
            int rows2 = Range2.GetLength(0);
            int cols2 = Range2.GetLength(1);

            int rows = rows1 + rows2;

            object[,] united;
            if (cols1 != cols2) return null;
            string[] header = new string[cols1];
            if (Headers)
            {
                for (int j = 0; j < cols1; j++)
                    if (Range1[0, j].ToString() != Range2[0, j].ToString()) return null;
                for (int j = 0; j < cols1; j++)
                    header[j] = Range1[0, j].ToString();
                rows--;
            }
            united = new object[rows, cols1];
            int b = 0;
            for (int i = 0; i < rows1; i++)
            {
                for (int j = 0; j < cols1; j++)
                    united[b, j] = Range1[i, j];
                b++;
            }

            for (int i = Headers ? 1 : 0; i < rows2; i++)
            {
                for (int j = 0; j < cols1; j++)
                    united[b, j] = Range2[i, j];
                b++;
            }

            for (int i = Headers ? 1 : 0; i < united.GetLength(0); i++)
                for (int k = i + 1; k < united.GetLength(0); k++)
                    if (COMPARE_ROWS(united, i, united, k))
                        united = TRIM_ARRAY(k, united);
            return POSTPROCCESS(united, ref Headers);
        }

        [ExcelFunction(Description = "Returns a relational intersection of two sets as a 2d array.")]
        public static object[,] REL_INTERSECT([ExcelArgument("Required. First range for intersetion.")] object[,] Range1,
            [ExcelArgument("Required. Second range for intersetion.")] object[,] Range2,
            [ExcelArgument("Optional. Enter TRUE if your ranges have titles, otherwise - FALSE. Default value is TRUE.")] object Headers_arg)
        {
            bool Headers = Optional.Check(Headers_arg, true);
            bool Header1 = Headers;
            bool Header2 = Headers;
            Range1 = PREPROCCESS(Range1, ref Header1);
            Range2 = PREPROCCESS(Range2, ref Header2);
            if (Header1 ^ Header2 == true)
            {
                if (Header1) Range1 = TRIM_ARRAY(0, Range1);
                if (Header2) Range2 = TRIM_ARRAY(0, Range2);
            }
            Headers = Header1 && Header2;

            int rows1 = Range1.GetLength(0);
            int cols1 = Range1.GetLength(1);
            int rows2 = Range2.GetLength(0);
            int cols2 = Range2.GetLength(1);

            int rows = rows1 + rows2;

            if (cols1 != cols2) return null;

            if (Headers)
            {
                for (int j = 0; j < cols1; j++)
                    if (Range1[0, j].ToString() != Range2[0, j].ToString()) return null;
            }
            object[,] inter = new object[rows, cols1];
            int b = 0;

            for (int h = 0; h < rows1; h++)
                for (int i = 0; i < rows2; i++)
                {
                    if (COMPARE_ROWS(Range1, h, Range2, i))
                    {
                        for (int j = 0; j < cols1; j++)
                            inter[b, j] = Range2[i, j];
                        b++;
                    }
                }
            object[,] res = new object[b, cols1];

            for (int i = 0; i < b; i++)
            {
                for (int j = 0; j < cols1; j++)
                    res[i, j] = inter[i, j];
            }
            return POSTPROCCESS(res, ref Headers);
        }

        [ExcelFunction(Description = "Returns a relational substraction of two sets as a 2d array.")]
        public static object[,] REL_SUBSTRACT([ExcelArgument("Required. First range for substraction.")] object[,] Range1,
            [ExcelArgument("Required. Second range for substraction.")] object[,] Range2,
            [ExcelArgument("Optional. Enter TRUE if your ranges have titles, otherwise - FALSE. Default value is TRUE.")] object Headers_arg)
        {
            bool Headers = Optional.Check(Headers_arg, true);
            bool Header1 = Headers;
            bool Header2 = Headers;
            Range1 = PREPROCCESS(Range1, ref Header1);
            Range2 = PREPROCCESS(Range2, ref Header2);
            if (Header1 ^ Header2 == true)
            {
                if (Header1) Range1 = TRIM_ARRAY(0, Range1);
                if (Header2) Range2 = TRIM_ARRAY(0, Range2);
            }
            Headers = Header1 && Header2;

            int cols1 = Range1.GetLength(1);
            int cols2 = Range2.GetLength(1);

            if (cols1 != cols2) return null;

            string[] header = new string[cols1];
            if (Headers)
            {
                for (int j = 0; j < cols1; j++)
                    if (Range1[0, j].ToString() != Range2[0, j].ToString()) return null;
                for (int j = 0; j < cols1; j++)
                    header[j] = Range1[0, j].ToString();
            }
            object[,] inter = REL_INTERSECT(Range1, Range2, Headers);


            for (int h = Headers ? 1 : 0; h < Range1.GetLength(0); h++)
                for (int i = Headers ? 1 : 0; i < inter.GetLength(0); i++)

                    if (COMPARE_ROWS(Range1, h, inter, i))
                        Range1 = TRIM_ARRAY(h, Range1);

            return POSTPROCCESS(Range1, ref Headers);
        }

        [ExcelFunction(Description = "Returns a relational times multiplication of two sets as a 2d array.")]
        public static object[,] REL_TIMES([ExcelArgument("Required. First range for times function.")] object[,] Range1,
            [ExcelArgument("Required. Second range for times function.")] object[,] Range2,
            [ExcelArgument("Optional. Enter TRUE if your ranges have titles, otherwise - FALSE. Default value is TRUE.")] object Headers_arg)
        {
            bool Headers = Optional.Check(Headers_arg, true);
            bool Header1 = Headers;
            bool Header2 = Headers;
            Range1 = PREPROCCESS(Range1, ref Header1);
            Range2 = PREPROCCESS(Range2, ref Header2);
            if (Header1 ^ Header2 == true)
            {
                if (Header1) Range1 = TRIM_ARRAY(0, Range1);
                if (Header2) Range2 = TRIM_ARRAY(0, Range2);
            }
            Headers = Header1 && Header2;

            int rows1 = Range1.GetLength(0);
            int cols1 = Range1.GetLength(1);
            int rows2 = Range2.GetLength(0);
            int cols2 = Range2.GetLength(1);

            int rows = rows1 * rows2;
            int cols = cols1 + cols2;

            string[] header = new string[cols];
            if (Headers)
            {
                for (int j = 0; j < cols1; j++)
                    header[j] = Range1[0, j].ToString();
                for (int j = 0; j < cols2; j++)
                    header[j + cols1] = Range2[0, j].ToString();
                rows = (rows1 - 1) * (rows2 - 1) + 1;
            }
            string[,] times = new string[rows, cols];

            if (Headers)
                for (int h = 0; h < cols; h++)
                    times[0, h] = header[h];

            int g = (Headers ? 1 : 0);
            for (int k = (Headers ? 1 : 0); k < rows1; k++)
            {
                for (int i = (Headers ? 1 : 0); i < rows2; i++)

                {
                    for (int j = 0; j < cols1; j++)
                        times[g, j] = Range1[k, j].ToString();
                    for (int j = 0; j < cols2; j++)
                        times[g, j + cols1] = Range2[i, j].ToString();
                    g++;
                }

            }

            return POSTPROCCESS(times, ref Headers);
        }

        [ExcelFunction(Description = "Returns a relational projection of the range as a 2d array.")]
        public static object[,] REL_PROJECTION([ExcelArgument("Required. The range for projection.")] object[,] Range,
            [ExcelArgument("Required. Column names (or column numbers starting with 0, if you have no headers), splited with coma, on which you want to project your range.")] object Column,
            [ExcelArgument("Optional. Enter TRUE if your ranges have titles, otherwise - FALSE. Default value is TRUE.")] object Headers_arg)
        {
            bool Headers = Optional.Check(Headers_arg, true);
            Range = PREPROCCESS(Range, ref Headers);


            int rows = Range.GetLength(0);
            int cols = Range.GetLength(1);

            string[] Columns = Column.ToString().Split(',').ToArray();
            int col_edit = Columns.Length;

            object[,] projection = new object[rows, col_edit];

            if (Headers)
            {
                int k = 0;
                for (int i = 0; i < cols; i++)
                {
                    if (Columns.Contains(Range[0, i].ToString()))
                    {
                        projection[0, k] = Range[0, i];
                        k++;
                    }
                }
            }

            for (int i = (Headers ? 1 : 0); i < rows; i++)
            {
                int k = 0;
                for (int j = 0; j < cols; j++)
                {
                    if (Headers)
                    {
                        if (!Columns.Contains(Range[0, j].ToString()))
                            continue;
                    }
                    else
                        if (!Columns.Contains(j.ToString()))
                        continue;

                    projection[i, k] = Range[i, j];
                    k++;
                }
            }

            return POSTPROCCESS(projection, ref Headers);
        }

        [ExcelFunction(Description = "Returns an relational selection of a set as a 2d array.")]
        public static object[,] REL_SELECTION([ExcelArgument("Required. The range for selection.")] object[,] Range,
            [ExcelArgument("Required. Column name (or column number starting with 0, if you have no headers) of " +
            "your range to compare values of.")] object Column,
            [ExcelArgument("Required. The sign of comparision i.e. \"=\",\"<\", etc.")] string Compare,
            [ExcelArgument("Required. The value to compare your column values with.")] object CompareWith,
            [ExcelArgument("Optional. Enter TRUE if your ranges have titles, otherwise - FALSE. Default value is TRUE.")] object Headers_arg)
        {
            bool Headers = Optional.Check(Headers_arg, true);
            Range = PREPROCCESS(Range, ref Headers);


            int rows = Range.GetLength(0);
            int cols = Range.GetLength(1);

            int k = 0;
            int col_comparator = 0;
            if (Headers)
            {
                for (int j = 0; j < cols; j++)
                    if (Column.ToString() == Range[0, j].ToString())
                    { col_comparator = j; break; }
            }
            else col_comparator = int.Parse(Column.ToString());

            try
            {
                for (int i = Headers ? 1 : 0; i < rows; i++)
                {
                    if (MatchExpression(Range[i, col_comparator], Compare, CompareWith) == false)
                        Range[i, 0] = null;
                    else k++;
                }
            }
            catch
            {
                return new object[,] { { ExcelError.ExcelErrorValue } };
            }
            object[,] selection = new object[k + (Headers ? 1 : 0), cols];
            k = 0;
            for (int i = 0; i < rows; i++)
            {
                if (Range[i, 0] != null)
                {
                    for (int j = 0; j < cols; j++)
                        selection[k, j] = Range[i, j];
                    k++;
                }
            }
            return POSTPROCCESS(selection, ref Headers);
        }

        private static bool MatchExpression(object compare1, string sign, object compare2)
        {
            string s1 = compare1.ToString().Replace('.', ',');
            string s2 = compare2.ToString().Replace('.', ',');

            if (!(double.TryParse(s1, out _) || (double.TryParse(s2, out _))))
                switch (sign)
                {
                    case "=":
                        return compare1.ToString() == compare2.ToString();
                    case "<>":
                        return compare1.ToString() != compare2.ToString();
                    default:
                        throw new Exception("Wrong logical sign");
                }


            switch (sign)
            {
                case "<":
                    return double.Parse(s1) < double.Parse(s2);
                case "<=":
                    return double.Parse(s1) <= double.Parse(s2);
                case ">":
                    return double.Parse(s1) > double.Parse(s2);
                case ">=":
                    return double.Parse(s1) >= double.Parse(s2);
                case "=":
                    return double.Parse(s1) == double.Parse(s2);
                case "<>":
                    return double.Parse(s1) != double.Parse(s2);
                default:
                    throw new Exception("Wrong logical sign");
            }
        }

        [ExcelFunction(Description = "Returns an relational join (natural or theta) of two sets as a 2d array.")]
        public static object[,] REL_JOIN_NATURAL([ExcelArgument("Required. First range for joining.")] object[,] Range1,
              [ExcelArgument("Required. Second range for joining.")] object[,] Range2,
              [ExcelArgument("Required. Column name (or column number starting with 0, if you have no headers) of first range you want to compare with.")] object Column1,
              [ExcelArgument("Required. The sign of comparision i.e. \"=\",\"<\", etc.")] string Compare,
              [ExcelArgument("Required. Column name (or column number starting with 0, if you have no headers) of second range you want to compare with.")] object Column2,
              [ExcelArgument("Optional. Enter TRUE if your ranges have titles, otherwise - FALSE. Default value is TRUE.")] object Headers_arg)
        {
            bool Headers = Optional.Check(Headers_arg, true);
            Range1 = PREPROCCESS(Range1, ref Headers);
            Range2 = PREPROCCESS(Range2, ref Headers);

            int rows1 = Range1.GetLength(0);
            int cols1 = Range1.GetLength(1);
            int rows2 = Range2.GetLength(0);
            int cols2 = Range2.GetLength(1);

            int rows = rows1 * rows2;
            int cols = cols1 + cols2 - 1;

            int range1_col = 0, range2_col = 0;

            object[] header = new object[cols];
            if (Headers)
            {
                for (int j = 0; j < cols1; j++)
                {
                    header[j] = Range1[0, j];
                    if (Column1.ToString() == Range1[0, j].ToString()) range1_col = j;
                }
                int c = 0;
                for (int j = 0; j < cols2; j++)
                {
                    if (Column1.ToString() == Range2[0, j].ToString())
                        continue;
                    header[c + cols1] = Range2[0, j];
                    if (Column2.ToString() == Range2[0, j].ToString()) range2_col = j;
                    c++;
                }
            }
            else
            {
                range1_col = int.Parse(Column1.ToString());
                range2_col = int.Parse(Column2.ToString());
            }
            object[,] connect = new object[rows, cols];

            int b = 0;

            for (int h = Headers ? 1 : 0; h < rows1; h++)
                for (int i = Headers ? 1 : 0; i < rows2; i++)
                {
                    if (MatchExpression(Range1[h, range1_col], Compare, Range2[i, range2_col]))
                    {
                        for (int j = 0; j < cols1; j++)
                            connect[b, j] = Range1[h, j];

                        int r = 0;
                        for (int j = 0; j < cols2; j++)
                        {
                            if (range2_col == j) continue;
                            connect[b, r + cols1] = Range2[i, j];
                            r++;
                        }
                        b++;
                    }
                }
            object[,] res = new object[b + (Headers ? 1 : 0), cols];
            if (Headers)
                for (int i = 0; i < cols; i++)
                    res[0, i] = header[i];
            for (int i = (Headers ? 1 : 0); i < b + (Headers ? 1 : 0); i++)
                for (int j = 0; j < cols; j++)
                    res[i, j] = connect[i - (Headers ? 1 : 0), j];

            return POSTPROCCESS(res, ref Headers);

        }

        [ExcelFunction(Description = "Returns an relational join (natural or theta) of two sets as a 2d array.")]
        public static object[,] REL_JOIN_THETA([ExcelArgument("Required. First range for joining.")] object[,] Range1,
            [ExcelArgument("Required. Second range for joining.")] object[,] Range2,
            [ExcelArgument("Required. Column name (or column number starting with 0, if you have no headers) of first range you want to compare with.")] object Column1,
            [ExcelArgument("Required. The sign of comparision i.e. \"=\",\"<\", etc.")] string Compare,
            [ExcelArgument("Required. Column name (or column number starting with 0, if you have no headers) of second range you want to compare with.")] object Column2,
            [ExcelArgument("Optional. Enter TRUE if your ranges have titles, otherwise - FALSE. Default value is TRUE.")] object Headers_arg)
        {
            bool Headers = Optional.Check(Headers_arg, true);
            Range1 = PREPROCCESS(Range1, ref Headers);
            Range2 = PREPROCCESS(Range2, ref Headers);

            int rows1 = Range1.GetLength(0);
            int cols1 = Range1.GetLength(1);
            int rows2 = Range2.GetLength(0);
            int cols2 = Range2.GetLength(1);

            int rows = rows1 * rows2;
            int cols = cols1 + cols2;

            int range1_col = 0, range2_col = 0;

            object[] header = new object[cols];
            if (Headers)
            {
                for (int j = 0; j < cols1; j++)
                {
                    header[j] = Range1[0, j];
                    if (Column1.ToString() == Range1[0, j].ToString()) range1_col = j;
                }
                int c = 0;
                for (int j = 0; j < cols2; j++)
                {
                    header[c + cols1] = Range2[0, j];
                    if (Column2.ToString() == Range2[0, j].ToString()) range2_col = j;
                    c++;
                }
            }
            else
            {
                range1_col = int.Parse(Column1.ToString());
                range2_col = int.Parse(Column2.ToString());
            }
            object[,] connect = new object[rows, cols];

            int b = 0;

            for (int h = Headers ? 1 : 0; h < rows1; h++)
                for (int i = Headers ? 1 : 0; i < rows2; i++)
                {
                    if (MatchExpression(Range1[h, range1_col], Compare, Range2[i, range2_col]))
                    {
                        for (int j = 0; j < cols1; j++)
                            connect[b, j] = Range1[h, j];

                        int r = 0;
                        for (int j = 0; j < cols2; j++)
                        {
                            connect[b, r + cols1] = Range2[i, j];
                            r++;
                        }
                        b++;
                    }
                }
            object[,] res = new object[b + (Headers ? 1 : 0), cols];
            if (Headers)
                for (int i = 0; i < cols; i++)
                    res[0, i] = header[i];
            for (int i = (Headers ? 1 : 0); i < b + (Headers ? 1 : 0); i++)
                for (int j = 0; j < cols; j++)
                    res[i, j] = connect[i - (Headers ? 1 : 0), j];

            return POSTPROCCESS(res, ref Headers);

        }

        private static object[,] REMOVE_EMPTY(object[,] Range, ref bool Headers)
        {
            if (Headers)
                for (int i = 0; i < Range.GetLength(1); i++)
                    if (Range[0, i].ToString() == ExcelEmpty.Value.ToString())
                    {
                        Range = TRIM_ARRAY(0, Range);
                        Headers = false;
                        break;
                    }

            for (int i = Headers ? 1 : 0; i < Range.GetLength(0); i++)
                for (int j = 0; j < Range.GetLength(1); j++)
                    if (Range[i, j].ToString() == ExcelEmpty.Value.ToString())
                    {
                        Range = TRIM_ARRAY(i, Range);
                        break;
                    }

            return Range;
        }

        private static object[,] REMOVE_DUBLICATES(object[,] Range, ref bool Headers)
        {

            for (int i = Headers ? 1 : 0; i < Range.GetLength(0) - 1; i++)
                for (int h = i + 1; h < Range.GetLength(0); h++)
                    if (COMPARE_ROWS(Range, i, Range, h))
                        Range = TRIM_ARRAY(h--, Range);

            return Range;
        }

        private static object[,] TRIM_ARRAY(int rowToRemove, object[,] originalArray)
        {
            object[,] result = new object[originalArray.GetLength(0) - 1, originalArray.GetLength(1)];

            for (int i = 0, j = 0; i < originalArray.GetLength(0); i++)
            {
                if (i == rowToRemove)
                    continue;

                for (int k = 0; k < originalArray.GetLength(1); k++)
                    result[j, k] = originalArray[i, k];
                j++;
            }

            return result;
        }
        private static bool COMPARE_ROWS(object[,] Range1, int row1, object[,] Range2, int row2)
        {
            int cols = Range1.GetLength(1);
            for (int i = 0; i < cols; i++)
                if (Range1[row1, i].ToString() != Range2[row2, i].ToString()) return false;
            return true;
        }

        [ExcelFunction(Description = "Returns a relational binary division of two sets as a 2d array.")]
        public static object[,] REL_DIVISION_BINARY([ExcelArgument("Required. First range for division.")] object[,] Range1,
             [ExcelArgument("Required. Second range for division.")] object[,] Range2)
        {
            bool Headers = true;

            Range1 = PREPROCCESS(Range1, ref Headers);
            Range2 = PREPROCCESS(Range2, ref Headers);
            if (Headers == false) return null;

            int cols1 = Range1.GetLength(1);
            int cols2 = Range2.GetLength(1);

            int cols = cols1 - cols2;

            string headers = "";
            int c = 0;
            while (Range1[0, c].ToString() != Range2[0, 0].ToString())
            {
                headers += Range1[0, c].ToString() + ",";
                c++;
            }
            if (headers.Length > 0)
                headers = headers.Remove(headers.Length - 1);
            headers += ";";
            if (c != cols)
            {
                for (; c < cols; c++)
                    headers += Range1[0, c + Range2.GetLength(1)].ToString() + ",";
                if (headers.Length > 0)
                    headers = headers.Remove(headers.Length - 1);
            }


            var headers_l = headers.Split(';')[0];
            var headers_r = headers.Split(';')[1];

            if (headers_l.Length != 0 && headers_r.Length != 0)
            {
                object[,] timesl = REL_TIMES(REL_TIMES(REL_PROJECTION(Range1, headers_l, Headers), Range2, Headers), REL_PROJECTION(Range1, headers_r, Headers), Headers);
                object[,] subl = REL_PROJECTION(REL_SUBSTRACT(timesl, Range1, Headers), headers_l + "," + headers_r, Headers);
                object[,] resl = REL_SUBSTRACT(REL_PROJECTION(Range1, headers_l + "," + headers_r, Headers), subl, Headers);
                return POSTPROCCESS(resl, ref Headers);
            }
            else

            if (headers_l.Length == 0)
            {
                //right
                object[,] timesr = REL_TIMES(Range2, REL_PROJECTION(Range1, headers_r, Headers), Headers);
                object[,] subr = REL_PROJECTION(REL_SUBSTRACT(timesr, Range1, Headers), headers_r, Headers);
                object[,] resr = REL_SUBSTRACT(REL_PROJECTION(Range1, headers_r, Headers), subr, Headers);
                return POSTPROCCESS(resr, ref Headers);

            }
            else
            {
                //left
                object[,] timesl = REL_TIMES(REL_PROJECTION(Range1, headers_l, Headers), Range2, Headers);
                object[,] subl = REL_PROJECTION(REL_SUBSTRACT(timesl, Range1, Headers), headers_l, Headers);
                object[,] resl = REL_SUBSTRACT(REL_PROJECTION(Range1, headers_l, Headers), subl, Headers);
                return POSTPROCCESS(resl, ref Headers);

            }
        }

        [ExcelFunction(Description = "Returns a ternary division of two sets as a 2d array.")]
        public static object[,] REL_DIVISION_TERNARY([ExcelArgument("Required. First range for division.")] object[,] Range1,
             [ExcelArgument("Required. Second range for division.")] object[,] Range2,
             [ExcelArgument("Required. Linking range for division.")] object[,] Range)
        {
            bool Headers = true;

            Range1 = PREPROCCESS(Range1, ref Headers);
            Range2 = PREPROCCESS(Range2, ref Headers);
            Range = PREPROCCESS(Range, ref Headers);
            if (Headers == false) return null;

            string headers = "";
            for (int i = 0; i < Range1.GetLength(1) - 1; i++)
                headers += Range1[0, i].ToString() + ",";
            headers += Range1[0, Range1.GetLength(1) - 1].ToString();

            object[,] times = REL_TIMES(Range1, Range2, Headers);
            object[,] sub = REL_PROJECTION(REL_SUBSTRACT(times, Range, Headers), headers, Headers);
            object[,] res = REL_SUBSTRACT(Range1, sub, Headers);
            return POSTPROCCESS(res, ref Headers);

        }
    }

    internal static class Optional
    {
        internal static bool Check(object arg, bool defaultValue)
        {
            if (arg is bool)
                return bool.Parse(arg.ToString());

            return defaultValue;

        }
    }

}
