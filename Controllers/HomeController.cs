using ExcelHelperProject.Helpers;
using ExcelHelperProject.Models;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace ExcelHelperProject.Controllers
{
    
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        [Route("[controller]/DownloadExcel")]
        public IActionResult CreateExcelFile()
        {
            List<UserViewModel> list = new()
            {
                new UserViewModel{
                    Id = 1,
                    FirstName = "John",
                    LastName = "Doe",
                    Age = 29,
                    IsMarriage = false,
                    Salary = 50700.45M,
                    MemberNo = "01/2023",
                    Date = DateTime.Now
                },
                new UserViewModel
                {
                    Id = 1,
                    FirstName = "Jenny",
                    LastName = "Doe",
                    Age = 24,
                    IsMarriage = true,
                    Salary = 50800.45M,
                    MemberNo = "02/2023",
                    Date = DateTime.Now
                },

            };

            var byteArr = ExcelHelper<UserViewModel>.CreateExcelFile(list);
            return File(byteArr, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "users-file.xlsx");
        }
    }
}
