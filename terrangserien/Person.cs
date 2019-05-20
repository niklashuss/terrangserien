using Serilog;

namespace terrangserien
{
    public class Person
    {
        private Person(int id)
        {
            Id = id;
        }

        private Result[] result = new Result[6];

        public int Id { get; set; }
        public string Name { get; set; }
        public string Surname { get; set; }
        public string Gender { get; set; }
        public string SocialNumber { get; set; }
        public string Distance { get; set; }
        public string Number { get; set; }
        public string Klass { get; set; }

        public string Result0 { get { return result[0].ToString(); } }
        public string Result1 { get { return result[1].ToString(); } }
        public string Result2 { get { return result[2].ToString(); } }
        public string Result3 { get { return result[3].ToString(); } }
        public string Result4 { get { return result[4].ToString(); } }
        public string Result5 { get { return result[5].ToString(); } }

        public Result Result(int index)
        {
            return result[index];
        }

        public void Result(int index, Result result)
        {
            this.result[index] = result;
        }

        private static int id = 0;
        private static int CreateId()
        {
            return id++;
        }

        public static Person Create()
        {
            int id = CreateId();
            return new Person(id);
        }
    }
}
