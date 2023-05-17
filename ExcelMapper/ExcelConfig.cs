using System.Collections.Generic;
using System.Linq;
using ExcelMapper.CellOption;

namespace ExcelMapper
{
    /// <summary>
    ///     表格配置
    /// </summary>
    public class ExcelConfig<T, TCellConfig> where TCellConfig : BaseCellOption<T>
    {
        protected ExcelConfig()
        {
            FieldOption = new List<TCellConfig>();
            FieldCombox = new List<(string, string[])>();
        }

        /// <summary>
        ///     依据表头的设置
        /// </summary>
        protected IEnumerable<TCellConfig> FieldOption { get; set; }

        /// <summary>
        ///     依据表头设置下拉数据
        /// </summary>
        protected IEnumerable<(string, string[])> FieldCombox { get; set; }

        /// <summary>
        ///     表格数据
        /// </summary>
        protected IEnumerable<T> Data { get; set; }

        /// <summary>
        ///     获取表头
        /// </summary>
        protected IEnumerable<string> Header => FieldOption.Select(item => item.ExcelField);
        
        /// <summary>
        ///     添加普通单元格设置
        /// </summary>
        protected void Add(TCellConfig option)
        {
            FieldOption = FieldOption.Append(option);
        }

        /// <summary>
        ///     添加下拉设置
        /// </summary>
        protected void AddCombox(string fieldName, string[] fieldComboxData)
        {
            FieldCombox = FieldCombox.Append((fieldName, fieldComboxData));
        }
    }
}