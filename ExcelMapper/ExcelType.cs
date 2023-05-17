using System;
using System.IO;
using ExcelMapper.Impl;
using ExcelMapper.Interface;

namespace ExcelMapper
{
    internal static class ExcelType
    {
        internal static readonly Func<Stream, IExcelCellReader> CellReaderFunc = stream => new EpPlusCellReader(stream);
        internal static readonly Func<Stream, IExcelReader> ReaderFunc = stream => new EpPlusExcelReader(stream);
    }
}