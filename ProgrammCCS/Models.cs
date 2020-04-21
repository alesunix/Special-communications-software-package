using System;

namespace ProgramCCS
{
    public class Person
    {
        public static string Name { get; set; }
        public static string Pass { get; set; }
        public static string Access { get; set; }

        //Конструктор с двумя аргументами
        public Person(string name, string access)
        {
            Name = name;
            Access = access;
        }
    }

    public class Number
    {
        public static int Ns { get; set; }
        public static int Nn { get; set; }
        public static int Nr { get; set; }

        public static string Prefix_number { get; set; }
    }
}
