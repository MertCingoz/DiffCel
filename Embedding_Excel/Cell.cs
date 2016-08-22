using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EmbeddedExcel
{
    public class Cell
    {
        private string oldValue;
        private string newValue;
        private string adress;
        private string sheet;
        private string operation;
        public string OldValue 
        {
            get { return oldValue; }
            set { oldValue = value; }
        }
        public string NewValue
        {
            get { return newValue; }
            set { newValue = value; }
        }
        public string Adress
        {
            get { return adress; }
            set { adress = value; }
        }
        public string Sheet
        {
            get { return sheet; }
            set { sheet = value; }
        }
        public string Operation
        {
            get { return operation; }
            set { operation = value; }
        }
    }
}
