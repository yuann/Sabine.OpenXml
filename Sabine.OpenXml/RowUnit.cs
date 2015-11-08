using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Sabine.OpenXml
{
    public class RowUnit
    {
        public SheetUnit ParentSheetUnit { get; }
        public Row Row { get; }
        public int RowNumber { get { return (int)Row.RowIndex.Value; } }

        public IEnumerable<CellUnit> CellUnits
        {
            get
            {
                return Row.Select(x => new CellUnit(this, x as Cell));
            }
        }

        public RowUnit(SheetUnit parentSheetUnit, Row Row)
        {
            this.ParentSheetUnit = parentSheetUnit;
            this.Row = Row;
        }
    }
}
