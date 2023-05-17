using System;
using System.Drawing;
using ExcelMapper.Interface;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelMapper.Impl
{
    public class EpPlusCell : IReadCell<ExcelRangeBase>
    {
        public EpPlusCell(ExcelRangeBase excelCell)
        {
            Cell = excelCell;
            if (Cell.Start.Row == 1)
            {
                Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                Cell.Style.Fill.BackgroundColor.SetColor(Color.RoyalBlue);
                Cell.Style.Font.Color.SetColor(Color.White);
            }
        }

        public ExcelRangeBase Cell { get; set; }

        public int Row => Cell.Start.Row - 1;
        public int Col => Cell.Start.Column - 1;
        public string StringValue => Cell.Text.Trim();
        public Type ValueType => Cell.Value.GetType();

        public object Value
        {
            get => Cell.Value;
            set => Cell.Value = value;
        }

        public void CopyCellFrom(IReadCell cell)
        {
            if (cell is not IReadCell<ExcelRangeBase> tcell)
            {
                return;
            }
            Value = tcell.Cell.Value;
        }
    }
}