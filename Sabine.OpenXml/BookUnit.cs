using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Sabine.OpenXml
{
    public class BookUnit
    {
        public SharedStringTable SharedStringTable { get; }
        public WorkbookPart WorkbookPart { get; }

        public IEnumerable<SheetUnit> SheetUnits
        {
            get
            {
                return WorkbookPart.Workbook.Descendants<Sheet>()
                    .Select(sheet => new SheetUnit(this, sheet));
            }
        }

        public BookUnit(WorkbookPart wbPart)
        {
            this.WorkbookPart = wbPart;
            this.SharedStringTable = GetSharedStringTable(wbPart);
        }

        static SharedStringTable GetSharedStringTable(WorkbookPart wbPart)
        {
            var tablePart = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            return tablePart != null ? tablePart.SharedStringTable : null;
        }

    }
}
