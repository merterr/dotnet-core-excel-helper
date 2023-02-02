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

        public ExcelColumnNameAttribute(string name)
        {
            Name = name;
        }

        public string GetName()
        {
            return Name;
        }
    }
}
