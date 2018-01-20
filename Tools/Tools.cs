﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tools
{
    public class Tools
    {
        public static string DateToStringEn(DateTime date)
        {
            string[] monthList = { "", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
            string[] dayList = { "", "1st", "2nd", "3rd", "4th", "5th", "6th", "7th", "8th", "9th", "10th", "11th", "12th", "13th", "14th", "15th", "16th", "17th", "18th", "19th", "20th", "21st", "22nd", "23rd", "24th", "25th", "26th", "27th", "28th", "29th", "30th", "31st" };


            string year = date.Year.ToString();
            int month = date.Month;
            int day = date.Day;

            return string.Format(@"{0}.{1}.{2}", monthList[month], dayList[day], year);
        }
    }
}
