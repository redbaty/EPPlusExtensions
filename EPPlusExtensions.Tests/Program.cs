using System;
using System.IO;
using Bogus;
using EPPlusExtensions.Annotations;
using EPPlusExtensions.Extensions;

namespace EPPlusExtensions.Tests
{
    internal class Product
    {
        public int Id { get; set; }

        public double Price { get; set; }

        [ExcelColumn("Product Name")]
        public string Name { get; set; }

        public string Remove { get; set; }
    }

    internal class Program
    {
        public static void Main()
        {
            var people = new Faker<Product>()
                .RuleFor(i => i.Id, f => f.UniqueIndex)
                .RuleFor(i => i.Price, f => f.Random.Double(10, 1000))
                .RuleFor(i => i.Name, f => f.Company.CompanyName())
                .RuleFor(i => i.Remove, f => f.Hacker.Phrase());

            var generate = people.Generate(1000);
            var x = generate.CreateMapping()
                .AutoMap()
                .Property(i => i.Price, e =>
                {
                    e.TransformValue = o =>
                    {
                        if (o is double d)
                        {
                            return $"Price is: {d:C}";
                        }

                        return o;
                    };
                })
                .RemovePropertyMapping(i => i.Remove);
            File.WriteAllBytes("output.xlsx", x.WriteExcelFile());
        }
    }
}