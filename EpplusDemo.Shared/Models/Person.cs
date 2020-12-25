using System;

namespace EpplusDemo.Shared.Models
{
    public class Person
    {
        public Person(int id)
        {
            Id = id;
        }
        /// <summary>
        /// ID
        /// </summary>
        public int Id { get; private set; }
        /// <summary>
        /// First name
        /// </summary>
        public string FirstName { get; set; }
        /// <summary>
        /// Last name
        /// </summary>
        public string LastName { get; set; }
        /// <summary>
        /// Age
        /// </summary>
        public int Age { get; set; }
        /// <summary>
        /// Created date
        /// </summary>
        public DateTime Created { get; set; }
    }
}