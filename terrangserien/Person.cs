using Serilog;

namespace terrangserien
{
    public class Person
    {

        private Person(int id)
        {
            Id = id;
        }

        public int Id { get; set; }
        public string Name { get; set; }
        public string Surname { get; set; }
        public string Gender { get; set; }
        public string SocialNumber { get; set; }
        public string Distance { get; set; }
        public string Number { get; set; }
        public string Klass { get; set; }
        public string Result0 { get; set; }
        public string Result1 { get; set; }
        public string Result2 { get; set; }
        public string Result3 { get; set; }
        public string Result4 { get; set; }
        public string Result5 { get; set; }

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
