﻿using System;

namespace Список_на_курьерские_отправления
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
}
