using System;

namespace ExcelMapper.CellOption
{
    /// <summary>
    ///     导出单元格设置
    /// </summary>
    public class ExportCellOption<T> : BaseCellOption<T>
    {
        private Func<T, object> _action;

        /// <summary>
        ///     转换 表格 数据的方法
        /// </summary>
        public Func<T, object> Action
        {
            get
            {
                if (_action != null)
                {
                    return _action;
                }
                _action = item => item.GetValue(PropName);
                return _action;
            }
            set => _action = value;
        }
    }
}