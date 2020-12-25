using System.Reflection;

namespace EpplusDemo.Shared.CdExcel
{
    public class CdExcelMemberInfo
    {
        /// <summary>
        /// 成员信息
        /// </summary>
        public MemberInfo MemberInfo { get; set; }
        /// <summary>
        /// 属性信息
        /// </summary>
        public PropertyInfo PropertyInfo { get; set; }
        /// <summary>
        /// 列名
        /// </summary>
        public string ColumnName { get; set; }
        /// <summary>
        /// 排序
        /// </summary>
        public int Order { get; set; }
    }
}