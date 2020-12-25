using System.IO;
using System.Linq;
using EpplusDemo.Shared.CdExcel;
using EpplusDemo.Shared.Models;
using EpplusDemo.Shared.Services;

namespace EpplusDemo.Client
{
    /// <summary>
    /// 导出测试
    /// </summary>
    public class ExportTest
    {
        public static void ExportRawModel()
        {
            var personService = new PersonService();
            var list = personService.GetPersonList();
            var stream = list.ExportToExcel("个人信息");
            stream.Position = 0;
            var filename = @"D:\epplus-demo-raw-data.xlsx";
            using var fileStream = new FileStream(filename,FileMode.Create,FileAccess.Write);
            stream.CopyTo(fileStream);
        }
        
        public static void ExportMapperModel()
        {
            var personService = new PersonService();
            var list = personService.GetPersonList();
            var models = list.Select(x => new PersonExportModel
            {
                Age = x.Age,
                Created = x.Created.ToString("yyyy-MM-dd HH:mm:ss"),
                Id = x.Id,
                FirstName = x.FirstName,
                LastName = x.LastName
            }).ToList();
            var stream = models.ExportToExcel("个人信息");
            stream.Position = 0;
            var filename = @"D:\epplus-demo-mapper-data.xlsx";
            using var fileStream = new FileStream(filename,FileMode.Create,FileAccess.Write);
            stream.CopyTo(fileStream);
        }
    }
}