using System.Collections.Generic;
using System.IO;

namespace ExcelMapper.Interface
{
    public interface IExcelCellReader : IExcelContainer<IReadCell>
    {
        void SetCombox(int col, int endRow, IEnumerable<string> data);
        void Save(Dictionary<int, int> dictWidth, bool autofit);
        Stream GetStream(Dictionary<int, int> dictWidth, bool autofit);
    }
}