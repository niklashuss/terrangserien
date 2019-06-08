using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace terrangserien
{
    class LooseFunctions
    {
        public static string GetClassFromSocialNumber(string socialNumber)
        {
            int year = ExtractYearFromSocialNumber(socialNumber);
            if (year < 0)
            {
                return "";
            }
            return YearToAgeClass(year);
        }

        public static string YearToAgeClass(int year)
        {
            DateTime now = DateTime.Today;
            string fullYear = now.ToString("yyyy").Substring(2, 2);
            int apa = Int32.Parse(fullYear);
            int diff = apa - year;
            if (diff < 5)
            {
                return "0-4";
            }
            else if (diff < 7)
            {
                return "5-6";
            }
            else if (diff < 9)
            {
                return "7-8";
            }
            else if (diff < 11)
            {
                return "9-10";
            }
            else if (diff < 13)
            {
                return "11-12";
            }
            else if (diff < 15)
            {
                return "13-14";
            }
            else if (diff < 17)
            {
                return "15-16";
            }
            return "";
        }

        public static int ExtractYearFromSocialNumber(string socialNumber)
        {
            if (socialNumber.Length == 0)
            {
                return -1;
            }

            if (socialNumber.StartsWith("20") || socialNumber.StartsWith("19"))
            {
                string year = socialNumber.Substring(2, 2);
                return Int32.Parse(year);
            }
            {
                string year = socialNumber.Substring(0, 2);
                return Int32.Parse(year);
            }
        }
    }
}
