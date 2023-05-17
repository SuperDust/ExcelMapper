using System.Reflection;

namespace ExcelMapper.CellOption
{
    public class BaseCellOption<T>
    {
        private PropertyInfo _prop;
        private string _propName;

        /// <summary>
        /// 宽度
        /// </summary>
        public int  Width { get; set; }
        /// <summary>
        ///     对应excel中的表头字段
        /// </summary>
        public string ExcelField { get; set; }

        /// <summary>
        ///     对应字段的属性(实际上包含PropName)
        /// </summary>
        public virtual PropertyInfo Prop
        {
            get => _prop;
            set
            {
                _prop = value;
                _propName = _prop.Name;
            }
        }

        /// <summary>
        ///     就是一个看起来比较方便的标识
        /// </summary>
        public virtual string PropName
        {
            get => _propName;
            set
            {
                _propName = value;
                _prop = typeof(T).GetProperty(_propName);
            }
        }
    }
}