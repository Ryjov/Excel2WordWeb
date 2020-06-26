using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.EntityFrameworkCore;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using System.IO;
using Excel2WordWeb.Models;
using Spire.Doc;
using Spire.Xls;
using Spire.Doc.Documents;

namespace Excel2WordWeb.Controllers
{
    public class HomeController : Controller
    {
        ApplicationContext _context;
        IWebHostEnvironment _appEnvironment;

        public HomeController(ApplicationContext context, IWebHostEnvironment appEnvironment)
        {
            _context = context;
            _appEnvironment = appEnvironment;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> AddFile(IFormFileCollection uploadedFiles)
        {
            foreach (var uploadedFile in uploadedFiles)
            {
                // путь к папке Files
                string path = $@"\Files\{uploadedFile.FileName}";
                // сохраняем файл в папку Files в каталоге wwwroot
                using (var fileStream = new FileStream(_appEnvironment.WebRootPath + path, FileMode.Create))
                {
                    await uploadedFile.CopyToAsync(fileStream);
                }
                FileModel file = new FileModel { Name = uploadedFile.FileName, Path = path };
                _context.Files.Add(file);
            }
            FindAndReplaceObject obj =
                new FindAndReplaceObject ($@"{_appEnvironment.WebRootPath}\Files\{uploadedFiles[0].FileName}",
                    $@"{_appEnvironment.WebRootPath}\Files\{uploadedFiles[1].FileName}", _appEnvironment.WebRootPath );
            obj.FindAndReplace();
            _context.SaveChanges();
            return RedirectToAction("Index");
        }
    }
}
