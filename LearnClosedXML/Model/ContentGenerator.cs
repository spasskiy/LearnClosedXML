using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LearnClosedXML
{
    internal class ContentGenerator
    {
        private static readonly string[] Names = { "Alice", "Bob", "Charlie", "David", "Eve", "Frank", "Grace", "Heidi", "Ivan", "Julia" };
        private static readonly string[] Addresses = { "123 Main St", "456 Elm St", "789 Oak St", "321 Pine St", "654 Maple St" };

        public List<User> GenerateUsers()
        {
            var users = new List<User>();
            var random = new Random();

            for (int i = 0; i < 50; i++)
            {
                var user = new User
                {
                    Name = Names[random.Next(Names.Length)],
                    Age = random.Next(18, 65), 
                    Address = Addresses[random.Next(Addresses.Length)],
                    Status = random.Next(2) == 1,
                    Salary = Math.Round(random.NextDouble() * 100000 + 20000, 2)
                };
                users.Add(user);
            }
            return users;
        }
    }

    internal class User
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string Address { get; set; }
        public bool Status { get; set; }
        public double Salary { get; set; }
    }
}
