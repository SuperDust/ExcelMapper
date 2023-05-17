using System;

namespace ExcelMapper.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class ExcelAttribute : Attribute
    {
        /// <summary>
        ///     excel字段
        /// </summary>
        public string ExcelField { get; set; } = "name";

        /// <summary>
        ///     读取内容转表达式 (如: 0=男,1=女,2=未知)
        /// </summary>
        public string ReadConverterExp { get; set; } = null;

        /// <summary>
        ///     设置只能选择不能输入的列内容
        /// </summary>
        public string[] Combox { get; set; } = null;

        /// <summary>
        ///     分隔符，读取字符串组内容
        /// </summary>
        public string Separator { get; set; } = ",";

        /// <summary>
        ///  导出使用宽度
        /// </summary>
        public int Width { get; set; } = 0;
    }
}