using ExcelHelperProject.Attributes;
using System;

namespace ExcelHelperProject.Models
{
    public class UserViewModel
    {
        [ExcelColumnName("Id", 8)]
        public int Id { get; set; }

        [ExcelColumnName("Member No",1)]
        public string MemberNo { get; set; }

        [ExcelColumnName("FirstName",2)]
        public string FirstName { get; set; }

        [ExcelColumnName("LastName",3)]
        public string LastName { get; set; }

        [ExcelColumnName("Is Married",4)]
        public bool IsMarried { get; set; }

        [ExcelColumnName("Age",5)]
        public int Age { get; set; }

        [ExcelColumnName("Salary",6)]
        public decimal Salary { get; set; }

        [ExcelColumnName("Date",7)]
        public DateTime Date { get; set; }
        public bool IsMarriage { get; internal set; }
    }
}
