namespace EpplusDemo.Shared.CdExcel
{
    /// <summary>
    /// Excel的表头信息
    /// </summary>
    public class CdExcelHeader
    {
        /// <summary>
        /// 列头名称
        /// </summary>
        public string ColumnName { get; set; }
        /// <summary>
        /// 属性名称
        /// </summary>
        public string PropertyName { get; set; }
        /// <summary>
        /// 列的顺序
        /// </summary>
        public int Order { get; set; }
    }
}