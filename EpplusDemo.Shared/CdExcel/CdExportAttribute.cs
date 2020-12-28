using System;

namespace EpplusDemo.Shared.CdExcel
{
    /// <summary>
    /// 导出表格的特性类
    /// </summary>
    [AttributeUsage(AttributeTargets.All)]
    public class CdExportAttribute : Attribute
    {
        /// <summary>
        /// 列名
        /// </summary>
        public string ColumnName { get; set; }
        /// <summary>
        /// 是否忽略此属性
        /// </summary>
        public bool Ignore { get; set; }
        /// <summary>
        /// 列的顺序
        /// </summary>
        public int ColumnOrder { get; set; }
    }
}