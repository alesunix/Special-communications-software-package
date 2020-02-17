using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProgramCCS
{
    class ClassComboBox
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
}
