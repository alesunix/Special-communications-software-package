using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProgramCCS
{
    public class ClassComboBox
    {
        public readonly int Value;
        public readonly string Text;
        public ClassComboBox(int Value, string Text)
        {
            this.Value = Value;
            this.Text = Text;
        }
        public override string ToString()
        {
            return this.Text;
        }
    }
    public class ClassComboBoxOblast //Класс для списка областей
    {
        public readonly string Value;
        public readonly string Text;
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
}
