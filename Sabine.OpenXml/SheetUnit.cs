using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Sabine.OpenXml
{
    public class SheetUnit
    {
        public BookUnit ParentBookUnit { get; }
        public Sheet Sheet { get; }
        public WorksheetPart WorksheetPart { get; }
        public string Name { get { return Sheet.Name; } }

        public IEnumerable<RowUnit> RowUnits
        {
            get
            {
                return WorksheetPart != null
                    ? WorksheetPart.Worksheet.Descendants<Row>().Select(x => new RowUnit(this, x))
                    : Enumerable.Empty<RowUnit>();
            }
        }

        public SheetUnit(BookUnit parentBookUnit, Sheet sheet)
        {
            this.ParentBookUnit = parentBookUnit;
            this.Sheet = sheet;
            this.WorksheetPart = parentBookUnit.WorkbookPart.GetPartById(Sheet.Id) as WorksheetPart;
        }
    }
}
