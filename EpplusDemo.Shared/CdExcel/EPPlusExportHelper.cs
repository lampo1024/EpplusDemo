using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EpplusDemo.Shared.CdExcel
{
    /// <summary>
    /// EPPlus导出帮助扩展类
    /// </summary>
    public static class EPPlusExportHelper
    {
        /// <summary>
        /// 导出列表到Excel,返回文件流格式
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheetName"></param>
        /// <param name="list"></param>
        /// <returns></returns>
        public static Stream ExportToExcel<T>(this List<T> list, string sheetName) where T : class
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var pck = new ExcelPackage();
            //创建工作表
            var ws = pck.Workbook.Worksheets.Add(sheetName);

            var epcExcelInfo = new CdExcelInfo<T>(list);
            for (var i = 0; i < epcExcelInfo.Headers.Count(); i++)
            {
                var header = epcExcelInfo.Headers[i];
                ws.Cells[1, i + 1].Value = header.ColumnName;
            }

            //填充数据
            if (list.Count > 0)
            {
                ws.Cells["A2"].LoadFromCollection(list
                    , false
                    , OfficeOpenXml.Table.TableStyles.None
                    , BindingFlags.Instance | BindingFlags.Public
                    , epcExcelInfo.IncludeMembers.ToArray());
            }
            ws.Cells.AutoFitColumns();
            //格式化表头
            using (var rng = ws.Cells["A1:BZ1"])
            {
                rng.Style.Font.Bold = true;
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            }
            var stream = new MemoryStream();
            pck.SaveAs(stream);
            stream.Position = 0;
            return stream;
        }
        
        /// <summary
        /// 导出列表到Excel,返回文件流格式
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheetName"></param>
        /// <param name="list"></param>
        /// <returns></returns>
        public static Stream ExportToExcel<T>(this List<T> list) where T : class
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var pck = new ExcelPackage();
            //创建工作表
            var ws = pck.Workbook.Worksheets.Add("sheetName");
            
            //填充数据
            if (list.Count > 0)
            {
                ws.Cells["A1"].LoadFromCollection(list);
            }
            ws.Cells.AutoFitColumns();
            var stream = new MemoryStream();
            pck.SaveAs(stream);
            stream.Position = 0;
            return stream;
        }
        
        /// <summary>
        /// 从泛型集合中加载Excel数据集的扩展方法
        /// </summary>
        /// <param name="excelRange">EPPlus的ExcelRange实例</param>
        /// <param name="collection">泛型集合</param>
        /// <param name="includedMemberInfoList"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static ExcelRangeBase LoadFromCollectionFiltered<T>(this ExcelRangeBase excelRange, IEnumerable<T> collection, MemberInfo[] includedMemberInfoList) where T : class
        {
            return excelRange.LoadFromCollection<T>(collection, false,
                OfficeOpenXml.Table.TableStyles.None,
                BindingFlags.Instance | BindingFlags.Public,
                includedMemberInfoList.ToArray());
        }
        public static ExcelRangeBase LoadFromCollection<T>(this ExcelRangeBase excelRange,
            IEnumerable<T> collection, string[] propertyNames, bool printHeaders) where T : class
        {
            MemberInfo[] membersToInclude = typeof(T)
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Where(p => propertyNames.Contains(p.Name))
                .ToArray();

            return excelRange.LoadFromCollection<T>(collection, printHeaders,
                OfficeOpenXml.Table.TableStyles.None,
                BindingFlags.Instance | BindingFlags.Public,
                membersToInclude);
        }
    }
}