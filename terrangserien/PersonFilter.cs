using System;
using System.Collections.Generic;
using System.Linq;

namespace terrangserien
{
    public class PersonFilter
    {
        public string Name { get; set; }
        public string Surname { get; set; }
        public string Gender { get; set; }
        public string Number { get; set; }
        public string SocialNumber { get; set; }
        public string Distance { get; set; }
    }

    public static class PersonListExtensions
    {
        public static IList<Person> FilterPersons(this IList<Person> persons, PersonFilter filter)
        {
            var q = persons
                .Where(p => string.IsNullOrWhiteSpace(filter.Gender) || p.Gender.ToLower() == filter.Gender.ToLower().Trim())
                .Where(p => string.IsNullOrWhiteSpace(filter.Name) || p.Name.ToLower().StartsWith(filter.Name.ToLower().Trim()))
                .Where(p => string.IsNullOrWhiteSpace(filter.Surname) || p.Surname.ToLower().StartsWith(filter.Surname.ToLower().Trim()))
                .Where(p => string.IsNullOrWhiteSpace(filter.Number) || p.Number.StartsWith(filter.Number.Trim()))
                .Where(p => string.IsNullOrWhiteSpace(filter.SocialNumber) || p.SocialNumber.StartsWith(filter.SocialNumber.Trim()))
                .Where(p => string.IsNullOrWhiteSpace(filter.Distance) || p.Distance.StartsWith(filter.Distance.Trim()))
                ;
            return q.ToList();
        }
    }
}
