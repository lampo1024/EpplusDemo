using EpplusDemo.Shared.CdExcel;

namespace EpplusDemo.Shared.Models
{
    public class PersonExportModel
    {
        /// <summary>
        /// ID
        /// </summary>
        [CdExport(Ignore = true)]
        public int Id { get; set; }
        /// <summary>
        /// First name
        /// </summary>
        [CdExport(ColumnName = "名",ColumnOrder = 1)]
        public string FirstName { get; set; }
        /// <summary>
        /// Last name
        /// </summary>
        [CdExport(ColumnName = "姓",ColumnOrder = 0)]
        public string LastName { get; set; }
        /// <summary>
        /// Age
        /// </summary>
        [CdExport(ColumnName = "年龄",ColumnOrder = 2)]
        public int Age { get; set; }
        /// <summary>
        /// Created date
        /// </summary>
        [CdExport(ColumnName = "创建时间",ColumnOrder = 3)]
        public string Created { get; set; }
    }
}