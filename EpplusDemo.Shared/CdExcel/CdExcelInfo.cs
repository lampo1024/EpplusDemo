using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace EpplusDemo.Shared.CdExcel
{
    /// <summary>
    /// EPPlus导出信息初始化类
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class CdExcelInfo<T> where T : class
    {
        public CdExcelInfo(List<T> list)
        {
            IncludeMembers = new List<MemberInfo>();
            Headers = new List<CdExcelHeader>();
            Init(list);
        }
        /// <summary>
        /// 需要输出的实体对象的成员集合
        /// </summary>
        public List<MemberInfo> IncludeMembers { get; private set; }
        /// <summary>
        /// 表头列的信息集合
        /// </summary>
        public List<CdExcelHeader> Headers { get; private set; }

        /// <summary>
        /// 初始化表头和成员集合信息，调用此方法可以过滤掉(不导出)CdExportAttribute中忽略的属性
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        private void Init(List<T> list) 
        {
            var props = typeof(T).GetProperties(BindingFlags.Instance | BindingFlags.Public);
            IncludeMembers = new List<MemberInfo>();
            Headers = new List<CdExcelHeader>();
            var epcMemberInfoList = new List<CdExcelMemberInfo>();
            foreach (var pi in props)
            {
                var info = new CdExcelMemberInfo
                {
                    MemberInfo = pi,
                    PropertyInfo = pi,
                    Order = 0,
                    ColumnName = pi.Name
                };
                var attrs = pi.GetCustomAttributes(typeof(CdExportAttribute), true);
                if (attrs.Length <= 0)
                {
                    epcMemberInfoList.Add(info);
                    continue;
                }
                var attr = (CdExportAttribute)attrs.FirstOrDefault();
                if (attr == null)
                {
                    epcMemberInfoList.Add(info);
                    continue;
                }
                if (attr.Ignore)
                {
                    continue;
                }

                if (!string.IsNullOrEmpty(attr.ColumnName))
                {
                    info.ColumnName = attr.ColumnName;
                }
                info.Order = attr.ColumnOrder;
                epcMemberInfoList.Add(info);
            }
            IncludeMembers = epcMemberInfoList.OrderBy(x => x.Order).Select(x => x.MemberInfo).ToList();
            var headers = epcMemberInfoList
                .Select(m => new CdExcelHeader
                {
                    Order = m.Order,
                    PropertyName = m.PropertyInfo.Name,
                    ColumnName = m.ColumnName
                }).OrderBy(x => x.Order)
                .ToList();
            Headers.AddRange(headers);
        }
    }
}