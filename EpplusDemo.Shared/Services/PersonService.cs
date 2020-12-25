using System;
using System.Collections.Generic;
using System.Linq;
using Bogus;
using Bogus.DataSets;
using Person = EpplusDemo.Shared.Models.Person;

namespace EpplusDemo.Shared.Services
{
    /// <summary>
    /// 服务类
    /// </summary>
    public class PersonService
    {
        public List<Person> GetPersonList()
        {
            var list = Enumerable.Range(1, 5)
                .Select(PersonFake)
                .ToList();
            return list;
        }

        Person PersonFake(int seed)
        {
            var person = new Faker<Person>()
                .CustomInstantiator(f => new Person(f.Random.Short(1,1000)))
                .RuleFor(p => p.FirstName, (f, u) => f.Name.FirstName(Name.Gender.Male))
                .RuleFor(p => p.LastName, (f, u) => f.Name.LastName(Name.Gender.Male))
                .RuleFor(p => p.Age, (f, u) => f.Random.Byte(18,30))
                .RuleFor(p => p.Created, (f, u) => DateTime.Now)
                .Generate();
            return person;
        }
    }
}