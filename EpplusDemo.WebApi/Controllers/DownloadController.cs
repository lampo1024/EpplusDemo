using System.IO;
using System.Linq;
using EpplusDemo.Shared.CdExcel;
using EpplusDemo.Shared.Models;
using EpplusDemo.Shared.Services;
using Microsoft.AspNetCore.Mvc;

namespace EpplusDemo.WebApi.Controllers
{
    public class DownloadController : ControllerBase
    {
        /// <summary>
        /// 不使用扩展方法的原始导出方法
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("download/export_raw_excel")]
        public IActionResult ExportRawExcel()
        {
            var personService = new PersonService();
            var data = personService.GetPersonList();
            var stream = EPPlusExportHelper.ExportToExcel(data, "个人信息");
            var contentType = "application/octet-stream";
            var fileName = "个人信息表.xlsx";
            return File(stream, contentType, fileName);
        }
        
        /// <summary>
        /// 使用扩展方法的自定义导出
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("download/export_mapper_excel")]
        public IActionResult ExportExcel()
        {
            var personService = new PersonService();
            var data = personService.GetPersonList();
            var models = data.Select(x => new PersonExportModel
            {
                Age = x.Age,
                Created = x.Created.ToString("yyyy-MM-dd HH:mm:ss"),
                Id = x.Id,
                FirstName = x.FirstName,
                LastName = x.LastName
            }).ToList();
            var stream = EPPlusExportHelper.ExportToExcel(models, "个人信息");
            var contentType = "application/octet-stream";
            var fileName = "个人信息表.xlsx";
            return File(stream, contentType, fileName);
        }
    }
}