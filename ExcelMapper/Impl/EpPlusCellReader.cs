using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelMapper.Interface;
using OfficeOpenXml;

namespace ExcelMapper.Impl
{
    /// <summary>
    ///     使用EPPlus获取excel的单元格
    /// </summary>
    public class EpPlusCellReader : IExcelCellReader
    {
        private ExcelPackage _pack;
        private int _rowCount;
        private ExcelWorksheet _sheet;
        private Stream _stream;

        public EpPlusCellReader(Stream stream)
        {
            Init(stream);
        }

        public IEnumerable<string> Headers { get; private set; }

        public long RowCount => _rowCount;

        public IDictionary<string, int> HeadersWithIndex { get; private set; }

        public IEnumerable<IReadCell> this[long row] => _sheet.Cells[(int)row + 1, 1, (int)row + 1, Headers.Count()]
            .Select(item => new EpPlusCell(item));

        public IReadCell this[long row, long col] => new EpPlusCell(_sheet.Cells[(int)row + 1, (int)col + 1]);

        public void Dispose()
        {
            _stream.Dispose();
            _pack.Dispose();
        }

        public void SetCombox(int col, int endRow, IEnumerable<string> data)
        {
            var range = _sheet.Cells[1 + 1, 1 + col, endRow, 1 + col];
            var dropDown = _sheet.DataValidations.AddListValidation(range.Address);
            foreach (var value in data)
            {
                dropDown.Formula.Values.Add(value);
            }
        }

        public void Save(Dictionary<int, int> dictWidth, bool autofit)
        {
            SaveTo(_stream, dictWidth,autofit);
        }

        /// <summary>
        ///     获取流
        /// </summary>
        public Stream GetStream(Dictionary<int, int> dictWidth, bool autofit)
        {
            var ms = new MemoryStream();
            SaveTo(ms,dictWidth,autofit);
            return ms;
        }

        public IEnumerator<IEnumerable<IReadCell>> GetEnumerator()
        {
            for (var row = 0; row < _rowCount; row++)
            {
                yield return this[row];
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        /// <summary>
        ///     保存到流
        /// </summary>
        private void SaveTo(Stream stream, Dictionary<int, int> dictWidth, bool autofit = false)
        {
            if (autofit)
            {
                AutoSize();
            }
            foreach (var width in dictWidth.Where(width => width.Value>0))
            {
                _sheet.Column(width.Key+1).Width = width.Value;
            }
            stream.Seek(0, SeekOrigin.Begin);
            stream.SetLength(0);
            _pack.SaveAs(stream);
            stream.Seek(0, SeekOrigin.Begin);
        }

        private void Init(Stream stream)
        {
            stream.Seek(0, SeekOrigin.Begin);
            // 使用传入的流, 可在 Save 时修改/覆盖
            _stream = stream;
            Init();
        }


        private void Init()
        {
            _pack = new ExcelPackage();
            _sheet = _pack.Workbook.Worksheets.Add("sheet1");
            _rowCount = 0;
        }

        private void AutoSize()
        {
            _sheet.Cells[1, 1, 1, _sheet.Dimension.Columns].AutoFilter = true;
            _sheet.Cells[1, 1, 1, _sheet.Dimension.Columns].AutoFitColumns();
        }
    }
}