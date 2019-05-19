using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace terrangserien.test
{
    [TestClass]
    public class PersonListExtensionsGenderTest
    {
        private List<Person> _testData;

        [TestInitialize]
        public void TestInitialize()
        {
            _testData = new List<Person>
            {
                new Person() { Gender = "f" },
                new Person() { Gender = "p" },
                new Person() { Gender = "f" },
                new Person() { Gender = "f" },
                new Person() { Gender = "F" },
                new Person() { Gender = "p" },
                new Person() { Gender = "" },
            };
        }

        [TestMethod]
        public void PersonListExtensionsTests_GenderFilter_Ok()
        {
            PersonFilter filter = new PersonFilter{ Gender = "f" };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(4, Result.Count);
        }

        [TestMethod]
        public void PersonListExtensionsTests_GenderFilter_UsingCapital_Ok()
        {
            PersonFilter filter = new PersonFilter { Gender = "F" };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(4, Result.Count);
        }

        [TestMethod]
        public void PersonListExtensionsTests_GenderFilter_EmptyReturnsAll_Ok()
        {
            PersonFilter filter = new PersonFilter { Gender = "" };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(_testData.Count, Result.Count);
        }

        [TestMethod]
        public void PersonListExtensionsTests_GenderFilter_StartsWithEndsWithSpace_Ok()
        {
            PersonFilter filter = new PersonFilter { Gender = " p  " };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(2, Result.Count);
        }
    }

    [TestClass]
    public class PersonListExtensionsNameTest
    {
        private List<Person> _testData;
 
        [TestInitialize]
        public void TestInitialize()
        {
            _testData = new List<Person>
            {
                new Person() { Name = "Agnes" },
                new Person() { Name = "Thomas" },
                new Person() { Name = "Ann-Sofie" },
                new Person() { Name = "Anna" },
                new Person() { Name = "Siv" },
                new Person() { Name = "Sven" },
                new Person() { Name = "Lisen" },
                new Person() { Name = "Anna" },
                new Person() { Name = "Sofie" },
                new Person() { Name = "Klara" },
            };
        }

        [TestMethod]
        public void PersonListExtensionsTests_NameFilter_StartsWith_Ok()
        {
            PersonFilter filter = new PersonFilter { Name = "ann" };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(3, Result.Count);
        }

        [TestMethod]
        public void PersonListExtensionsTests_NameFilter_StartsWithUsingCapitals_Ok()
        {
            PersonFilter filter = new PersonFilter { Name = "AnN" };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(3, Result.Count);
        }

        [TestMethod]
        public void PersonListExtensionsTests_NameFilter_Empty_Ok()
        {
            PersonFilter filter = new PersonFilter { Name = "" };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(10, Result.Count);
        }

        [TestMethod]
        public void PersonListExtensionsTests_NameFilter_StartEndsWithSpace_Ok()
        {
            PersonFilter filter = new PersonFilter { Name = "  Klara  " };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(1, Result.Count);
        }

        [TestMethod]
        public void PersonListExtensionsTests_NameFilter_NotFound_Ok()
        {
            PersonFilter filter = new PersonFilter { Name = "Tine" };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(0, Result.Count);
        }
    }

    [TestClass]
    public class PersonListExtensionsSurnameTest
    {
        private List<Person> _testData;

        [TestInitialize]
        public void TestInitialize()
        {
            _testData = new List<Person>
            {
                new Person() { Surname = "Svensson" },
                new Person() { Surname = "Andersson" },
                new Person() { Surname = "Kristersson" },
                new Person() { Surname = "Dottirsson" },
                new Person() { Surname = "Ulgsson" },
                new Person() { Surname = "Andreasson" },
                new Person() { Surname = "Liamsson" },
                new Person() { Surname = "Gundersson" },
                new Person() { Surname = "Tillerud" },
                new Person() { Surname = "Andén" },
            };
        }

        [TestMethod]
        public void PersonListExtensionsTests_SurnameFilter_StartsWith_Ok()
        {
            PersonFilter filter = new PersonFilter { Surname = "And" };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(3, Result.Count);
        }

        [TestMethod]
        public void PersonListExtensionsTests_SurnameFilter_StartsWithUsingCapitals_Ok()
        {
            PersonFilter filter = new PersonFilter { Surname = "AnD" };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(3, Result.Count);
        }

        [TestMethod]
        public void PersonListExtensionsTests_SurnameFilter_Empty_Ok()
        {
            PersonFilter filter = new PersonFilter { Surname = "" };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(10, Result.Count);
        }

        [TestMethod]
        public void PersonListExtensionsTests_SurnameFilter_StartEndsWithSpace_Ok()
        {
            PersonFilter filter = new PersonFilter { Surname = "  Tillerud " };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(1, Result.Count);
        }

        [TestMethod]
        public void PersonListExtensionsTests_SurnameFilter_NotFound_Ok()
        {
            PersonFilter filter = new PersonFilter { Surname = "Arnesson" };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(0, Result.Count);
        }
    }

    [TestClass]
    public class PersonListExtensionsNumberTest
    {
        private List<Person> _testData;

        [TestInitialize]
        public void TestInitialize()
        {
            _testData = new List<Person>
            {
                new Person() { Number = "1" },
                new Person() { Number = "103" },
                new Person() { Number = "423" },
                new Person() { Number = "3" },
                new Person() { Number = "111" },
                new Person() { Number = "112" },
                new Person() { Number = "113" },
                new Person() { Number = "114" },
            };
        }

        [TestMethod]
        public void PersonListExtensionsTests_NumberFilter_Ok()
        {
            PersonFilter filter = new PersonFilter { Number = "103" };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(1, Result.Count);
        }

        [TestMethod]
        public void PersonListExtensionsTests_NumberFilter_StartsWith_Ok()
        {
            PersonFilter filter = new PersonFilter { Number = "1" };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(6, Result.Count);
        }

        [TestMethod]
        public void PersonListExtensionsTests_NumberFilter_StartsWithEndsWithSpace_Ok()
        {
            PersonFilter filter = new PersonFilter { Number = "  3  " };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(1, Result.Count);
        }

        [TestMethod]
        public void PersonListExtensionsTests_NumberFilter_NotFound_Ok()
        {
            PersonFilter filter = new PersonFilter { Number = "512" };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(0, Result.Count);
        }
    }

    [TestClass]
    public class PersonListExtensionsSocialNumberTest
    {
        private List<Person> _testData;

        [TestInitialize]
        public void TestInitialize()
        {
            _testData = new List<Person>
            {
                new Person() { SocialNumber = "14" },
                new Person() { SocialNumber = "2014" },
                new Person() { SocialNumber = "2013" },
                new Person() { SocialNumber = "20140102-1212" },
                new Person() { SocialNumber = "130341-1423" },
                new Person() { SocialNumber = "20141202-1142" },
                new Person() { SocialNumber = "15" },
                new Person() { SocialNumber = "2014" },
                new Person() { SocialNumber = "2013" },
            };
        }

        [TestMethod]
        public void PersonListExtensionsTests_SocialNumberFilter_Ok()
        {
            PersonFilter filter = new PersonFilter { SocialNumber = "14" };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(1, Result.Count);
        }

        [TestMethod]
        public void PersonListExtensionsTests_SocialNumberFilter_StartsWith_Ok()
        {
            PersonFilter filter = new PersonFilter { SocialNumber = "2014" };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(4, Result.Count);
        }

        [TestMethod]
        public void PersonListExtensionsTests_SocialNumberFilter_StartsWithEndsWithSpace_Ok()
        {
            PersonFilter filter = new PersonFilter { SocialNumber = "  20141202-1142  " };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(1, Result.Count);
        }

        [TestMethod]
        public void PersonListExtensionsTests_SocialNumberFilter_NotFound_Ok()
        {
            PersonFilter filter = new PersonFilter { SocialNumber = "2011" };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(0, Result.Count);
        }
    }

    [TestClass]
    public class PersonListExtensionsDistanceTest
    {
        private List<Person> _testData;

        [TestInitialize]
        public void TestInitialize()
        {
            _testData = new List<Person>
            {
                new Person() { Distance = "600" },
                new Person() { Distance = "1200" },
                new Person() { Distance = "1200" },
                new Person() { Distance = "3000" },
                new Person() { Distance = "1300" },
                new Person() { Distance = "600" },
                new Person() { Distance = "600" },
                new Person() { Distance = "3000" },
                new Person() { Distance = "1200" },
                new Person() { Distance = "3000" },
            };
        }
        [TestMethod]
        public void PersonListExtensionsTests_DistanceFilter_Ok()
        {
            PersonFilter filter = new PersonFilter { Distance = "600" };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(3, Result.Count);
        }

        [TestMethod]
        public void PersonListExtensionsTests_DistanceFilter_StartsWith_Ok()
        {
            PersonFilter filter = new PersonFilter { Distance = "1" };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(4, Result.Count);
        }

        [TestMethod]
        public void PersonListExtensionsTests_DistanceFilter_StartsWithEndsWithSpace_Ok()
        {
            PersonFilter filter = new PersonFilter { Distance = "  1300  " };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(1, Result.Count);
        }

        [TestMethod]
        public void PersonListExtensionsTests_DistanceFilter_NotFound_Ok()
        {
            PersonFilter filter = new PersonFilter { Distance = "1400" };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(0, Result.Count);
        }
    }

    [TestClass]
    public class PersonListExtensionsAllFilterTest
    {
        private List<Person> _testData;

        [TestInitialize]
        public void TestInitialize()
        {
            _testData = new List<Person>
            {
                new Person() { Name = "Lisa", Surname = "Lo", Gender = "f", SocialNumber = "2014", Distance = "600" },
                new Person() { Name = "Arne", Surname = "Gustavsson", Gender = "p", SocialNumber = "20131231-1214", Distance = "600" },
                new Person() { Name = "Torbjörn", Surname = "Irsson", Gender = "p", SocialNumber = "2015", Distance = "1200" },
                new Person() { Name = "Jens", Surname = "Gustavsson", Gender = "p", SocialNumber = "2013", Distance = "1200" },
                new Person() { Name = "Sofie", Surname = "Olsson", Gender = "f", SocialNumber = "2014", Distance = "600" },
                new Person() { Name = "Hannes", Surname = "Lerman", Gender = "p", SocialNumber = "20130112", Distance = "1200" },
                new Person() { Name = "Ann-Katrin", Surname = "Brorsson", Gender = "f", SocialNumber = "20121231-5434", Distance = "600" },
                new Person() { Name = "Katrin", Surname = "Ekemur", Gender = "f", SocialNumber = "2014", Distance = "1200" },
                new Person() { Name = "Lisa", Surname = "Gustafsson", Gender = "f", SocialNumber = "2016", Distance = "1200" },
                new Person() { Name = "Laura", Surname = "Andersson", Gender = "f", SocialNumber = "20140505", Distance = "600" },
                new Person() { Name = "Jonas", Surname = "Grahn", Gender = "p", SocialNumber = "2013", Distance = "1200" },
            };
        }

        [TestMethod]
        public void PersonListExtensionsTests_DistanceGender_Ok()
        {
            PersonFilter filter = new PersonFilter { Distance = "600", Gender = "f" };
            var Result = _testData.FilterPersons(filter);
            Assert.AreEqual(4, Result.Count);
        }
    }
}

