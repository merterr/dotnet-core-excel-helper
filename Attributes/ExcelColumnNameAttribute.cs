using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelHelperProject.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExcelColumnNameAttribute : Attribute
    {
        private string Name { get; set; }

        private int Order { get; set; }

        public ExcelColumnNameAttribute(string name, int order)
        {
            Name = name;
            Order = order;
        }

        public string GetName()
        {
            return Name;
        }

        public int GetOrder()
        {
            return Order;
        }
    }
}
