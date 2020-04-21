using System;
using System.Data.SqlClient;

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

    public class Connection
    {
        public static SqlConnection con = new SqlConnection(@"Data Source=192.168.0.3;Initial Catalog=ccsbase;Persist Security Info=True;User ID=Lan;Password=Samsung0");
    }
}
