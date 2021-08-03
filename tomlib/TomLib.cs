using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace tomlib
{
    public class TomLib
    {
        public static string getWeek(DateTime dt)
        {
            string str = "";

            switch (dt.DayOfWeek)
            {
                case DayOfWeek.Monday:
                    str = "一";
                    break;
                case DayOfWeek.Tuesday:
                    str = "二";
                    break;
                case DayOfWeek.Wednesday:
                    str = "三";
                    break;
                case DayOfWeek.Thursday:
                    str = "四";
                    break;
                case DayOfWeek.Friday:
                    str = "五";
                    break;
                case DayOfWeek.Saturday:
                    str = "六";
                    break;
                case DayOfWeek.Sunday:
                    str = "日";
                    break;
            }

            return str;
        }

    }
}
