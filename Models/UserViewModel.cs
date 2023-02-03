using ExcelHelperProject.Attributes;
using System;

namespace ExcelHelperProject.Models
{
    public class UserViewModel
    {
        public int Id { get; set; }

        [ExcelColumnName("Member No")]
        public string MemberNo { get; set; }

        [ExcelColumnName("FirstName")]
        public string FirstName { get; set; }

        [ExcelColumnName("LastName")]
        public string LastName { get; set; }

        [ExcelColumnName("Is Married")]
        public bool IsMarried { get; set; }

        [ExcelColumnName("Age")]
        public int Age { get; set; }

        [ExcelColumnName("Salary")]
        public decimal Salary { get; set; }

        [ExcelColumnName("Date")]
        public DateTime Date { get; set; }


    }
}
