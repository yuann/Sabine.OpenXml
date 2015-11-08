using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Sabine.OpenXml
{
    public class CellUnit
    {
        public RowUnit ParentRowUnit { get; }
        public Cell Cell { get; }
        public int RowNumber { get { return ParentRowUnit.RowNumber; } }
        public int ColumnNumber { get { return 0; } }

        public CellUnit(RowUnit parentRowUnit, Cell cell)
        {
            this.ParentRowUnit = parentRowUnit;
            this.Cell = cell;
        }

        public object GetObject()
        {
            if (Cell.DataType.HasValue)
            {
                switch (Cell.DataType.Value)
                {
                    case CellValues.Boolean: return bool.Parse(Cell.InnerText);
                    case CellValues.Date: return DateTime.Parse(Cell.InnerText);
                    case CellValues.Number: return double.Parse(Cell.InnerText);
                    case CellValues.SharedString: return GetSharedString();
                }
            }

            return Cell.InnerText;
        }

        public string GetString()
        {
            if (Cell.DataType.HasValue && Cell.DataType.Value == CellValues.SharedString)
            {
                return GetSharedString();
            }

            return Cell.InnerText;
        }

        string GetSharedString()
        {
            var stringTable = ParentRowUnit.ParentSheetUnit.ParentBookUnit.SharedStringTable;
            if (stringTable != null)
            {
                return stringTable.ElementAt(int.Parse(Cell.InnerText)).InnerText;
            }

            return Cell.InnerText;
        }

        public override string ToString()
        {
            return GetString();
        }
    }
}
