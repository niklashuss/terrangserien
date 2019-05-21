using Serilog;
using System;

namespace terrangserien
{
    public class Result : IComparable
    {
        public double Time { get; set; }

        private Result(double value)
        {
            Time = value;
        }

        public override string ToString()
        {
            if (Time == 0.0)
            {
                return "";
            }
            return string.Format("{0:0.00}", Time);
        }

        public int CompareTo(object obj)
        {
            if (obj == null)
            {
                return 1;
            }
            Result other = obj as Result;
            if (other.Time > Time)
            {
                return -1;
            }
            if (other.Time == Time)
            {
                return 0;
            }
            return 1;
        }

        public static Result Create(double value)
        {
            return new Result(value);
        }

        public static Result Create(string value)
        {
            double dbl = 0.0f;
            if (value.Length > 0)
            {
                dbl = double.Parse(value);
            }
            return new Result(dbl);
        }
    }
}
