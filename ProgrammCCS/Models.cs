using System;
using System.Data;
using System.Data.SqlClient;

namespace ProgramCCS
{
    public class Connection
    {
        public static SqlConnection con = new SqlConnection(@"Data Source=192.168.0.3;Initial Catalog=ccsbase;Persist Security Info=True;User ID=Lan;Password=Samsung0");
    }

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

    public class ClassComboBoxOblast //Класс для списка областей
    {
        public readonly string Value;
        public readonly string Text;

        //Конструктор с двумя аргументами
        public ClassComboBoxOblast(string Value, string Text)
        {
            this.Value = Value;
            this.Text = Text;
        }
        public override string ToString()
        {
            return this.Text;
        }
    }

    public class Partner
    {
        public static string Name { get; set; }       
    }

    public class Dates
    {
        public static DateTime StartDate { get; set; }
        public static DateTime EndDate { get; set; }

        //Конструктор с двумя аргументами
        public Dates (DateTime startDate, DateTime endDate)
        {
            StartDate = startDate;
            EndDate = endDate;
        }
    }
    public class Table
    {
        //создаем экземпляр класса DataTable
        public static DataTable DtRegistry { get; set; }
        public static DataTable DtInvoice { get; set; }
        public static DataTable DtPartner { get; set; }       
    }

}
