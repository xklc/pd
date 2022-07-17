using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pd
{
    public class NumberConvertUtil
    {
        private const string initValue = "A0000001";

        private static string cs = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        public static string AddOne(string str)
        {
            if (str == string.Empty)
            {
                return "";
            }
            int pos = str.IndexOf('-');
            if (pos > 0)
            {
                Int64 num = StringToNumber(str.Substring(pos+1));
                return str.Substring(0,pos)+"-"+NumberToString(++num);
            }
            return "";
        }

        public static Int64 StringToNumber(string str)
        {
            int leg = str.Length;
            double num = 0;
            if (leg != 0)
            {
                for (int i = 0; i < leg; i++)
                {
                    if (str[i] != '0')
                    {
                        num += cs.IndexOf(str[i]) * Math.Pow(36, leg - 1 - i);
                    }
                }
            }
            return Convert.ToInt64(num);
        }

        public static string NumberToString(Int64 num)
        {
            string str = string.Empty;
            while (num >= 36)
            {
                str = cs[(int)(num % 36)] + str;
                num = num / 36;
            }
            return cs[(int)num] + str;
        }
    }
}
