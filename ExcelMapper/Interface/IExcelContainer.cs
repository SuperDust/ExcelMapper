using System;
using System.Collections.Generic;

namespace ExcelMapper.Interface
{
    public interface IExcelContainer<T> : IDisposable, IEnumerable<IEnumerable<T>>
    {
        /// <summary>
        ///     行数索引
        /// </summary>
        /// <param name="row">索引</param>
        IEnumerable<T> this[long row] { get; }

        /// <summary>
        ///     行列数索引
        /// </summary>
        /// <param name="row">行索引</param>
        /// <param name="col">列索引</param>
        T this[long row, long col] { get; }

        /// <summary>
        ///     表头索引 (读取excel)
        /// </summary>
        IDictionary<string, int> HeadersWithIndex { get; }
        
        /// <summary>
        ///     总行数 (读取excel)
        /// </summary>
        long RowCount { get; }
    }
}