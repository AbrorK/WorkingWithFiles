using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Hosting.Internal;
using Microsoft.AspNetCore.Mvc;
using WorkingWithFiles.enums;
using WorkingWithFiles.helpers;
using WorkingWithFiles.models;

namespace WorkingWithFiles.Controllers
{
    public class FileController : Controller
    {
        private readonly IHostingEnvironment _hostingEnvironment;

        public FileController(IHostingEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }
        public IActionResult Index()
        {
            return View();
        }

        [HttpGet("{id}")]
        public IActionResult DownloadDeal([FromRoute] int id)
        {
            var result = new LogisTicketItem();

            try
            {
                //result = _profileService.GetTicket(UserId, id);
            }
            catch
            {
                //return Error("No found deal");
            }

            var templateName = $"deal_{Enum.GetName(typeof(WarrantyType), result.warranty_id)}_template_qr.docx";
            var templatePath = Path.Combine(_hostingEnvironment.WebRootPath, "reports", templateName);
            var savePath = Path.Combine(_hostingEnvironment.WebRootPath, "tmp", $"{Guid.NewGuid()}_{templateName}");

            try
            {
                var tokenUrl = "test-url";
                result.DownloadDeal(templatePath, savePath, tokenUrl, _hostingEnvironment);

                var bytes = System.IO.File.ReadAllBytes(savePath);


                return File(bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", Path.GetFileName(savePath));
            }
            catch (Exception ex)
            {
                ///return Error(ex.Message);
            }
            finally
            {
                try
                {
                    System.IO.File.Delete(savePath);
                }
                catch
                {

                }
            }
        }

        //[ApiValidation(Order = 1)]
        //public virtual IActionResult Error(string message = "")
        //{
        //    if (Request.IsMobile())
        //    {
        //        if (string.IsNullOrWhiteSpace(message))
        //            message = "Произошла ошибка при обработке вашего запроса!";
        //        else
        //            message = PackageTools.CleanMessage(message);

        //        var response = new ApiResponse(message);

        //        return Ok(response);
        //    }
        //    else
        //    {
        //        throw new AppException(message);
        //    }

        //}
    }
}