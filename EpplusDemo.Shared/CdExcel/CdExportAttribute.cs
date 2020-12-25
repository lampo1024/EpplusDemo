using System;

namespace EpplusDemo.Shared.CdExcel
{
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