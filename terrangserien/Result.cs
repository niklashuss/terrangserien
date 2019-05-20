using System;

namespace terrangserien
{
    public class Result : IComparable
    {
        public int Minutes { get; set; }
        public int Seconds { get; set; }

        private Result(int minutes, int seconds)
        {
            Minutes = minutes;
            Seconds = seconds;
        }

        public override string ToString()
        {
            if (Minutes == 0 && Seconds == 0)
            {
                return "";
            }
            if (Seconds < 10)
            {
                return string.Format("{0},0{1}", Minutes, Seconds);
            }
            return string.Format("{0},{1}", Minutes, Seconds);
        }

        public int CompareTo(object obj)
        {
            if (obj == null)
            {
                return 1;
            }
            Result other = obj as Result;
            if (other.Minutes > Minutes)
            {
                return -1;
            }
            if (other.Minutes == Minutes)
            {
                if (other.Seconds == Seconds)
                {
                    return 0;
                }
                if (other.Seconds > Seconds)
                {
                    return -1;
                }
            }
            return 1;
        }

        public static Result Create(ref string str)
        {
            if (string.IsNullOrEmpty(str))
            {
                return new Result(0, 0);
            }

            string[] strings = str.Split(',');
            int minutes = int.Parse(strings[0]);
            if (strings.Length == 1)
            {
                return new Result(minutes, 0);
            }
            int seconds = int.Parse(strings[1]);
            return new Result(minutes, seconds);
        }
    }
}
