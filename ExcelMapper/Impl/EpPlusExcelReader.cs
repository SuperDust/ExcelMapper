using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelMapper.Interface;
using OfficeOpenXml;

namespace ExcelMapper.Impl
{
    public class EpPlusExcelReader : IExcelReader
    {
        private int _rowCount;
        private ExcelWorksheet _sheet;

        public EpPlusExcelReader(Stream stream)
        {
            var sheets = new ExcelPackage(stream).Workbook?.Worksheets;
            if (sheets is { Count: > 0 })
            {
                Init(sheets[0]);
            }
        }

        public IEnumerable<string> Headers { get; private set; }

        public IEnumerable<string> this[long row]
        {
            get
            {
                for (var i = 1; i <= Headers.Count(); i++)
                {
                    yield return _sheet.Cells[(int)row + 1, i, (int)row + 1, i].Value.ToString();
                }
            }
        }

        public string this[long row, long col] => _sheet.Cells[(int)row + 1, (int)col + 1].Value.ToString();

        public IDictionary<string, int> HeadersWithIndex { get; private set; }
      
        public long RowCount => _rowCount;

        public IEnumerator<IEnumerable<string>> GetEnumerator()
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

        public void Dispose()
        {
            _sheet.Workbook.Dispose();
        }

        private void Init(ExcelWorksheet sheet)
        {
            _sheet = sheet;
            // 清除空行
            for (var i = sheet.Dimension.End.Row; i >= sheet.Dimension.Start.Row; i--)
            {
                var isRowEmpty = true;
                for (var j = sheet.Dimension.Start.Column; j <= sheet.Dimension.End.Column; j++)
                {
                    if (sheet.Cells[i, j].Value == null)
                    {
                        continue;
                    }
                    isRowEmpty = false;
                    break;
                }
                if (isRowEmpty)
                {
                    sheet.DeleteRow(i);
                }
            }
            _rowCount = sheet.Dimension.Rows;
            HeadersWithIndex = sheet.Cells[1, 1, 1, sheet.Dimension.Columns]
                .Where(item => item.Value != null && !string.IsNullOrEmpty(item.Value.ToString()))
                .ToDictionary(item => item.Value.ToString().Trim(), item => item.End.Column - 1);
            Headers = HeadersWithIndex.Select(item => item.Key).ToList();
        }
    }
}