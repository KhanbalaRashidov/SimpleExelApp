using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using SimpleExelApp.Models;
using SimpleExelApp.Repositories;
using System.Drawing;

namespace SimpleExelApp.Controllers
{
    public class UsersController : Controller
    {
        private readonly IExcelRepository _excelRepository;

        public UsersController(IExcelRepository excelRepository)
        {
            _excelRepository = excelRepository;
        }

        
        public IActionResult Index()
        {
            var users =_excelRepository.GetUserList();

            return View(users);
        }

        public IActionResult ExportToExcel()
        {
            var stream =_excelRepository.ExportToExcel();
            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "users.xlsx");
        }

        [HttpGet]
        public IActionResult BatchUserUpload()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult BatchUserUpload(IFormFile batchUsers)
        {
            if (ModelState.IsValid)
            {
                if (batchUsers?.Length > 0)
                {
                    // convert to a stream
                    var stream = batchUsers.OpenReadStream();

                    List<User> users = new List<User>();

                    try
                    {
                       var data= _excelRepository.BatchUserUpload(batchUsers, stream, users);

                        return View("Index", data);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }

            return View();
        }

       
        
    }
}
