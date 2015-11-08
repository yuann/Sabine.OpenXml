using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;

namespace Sabine.OpenXml
{
    public static class SpreadsheetEx
    {
        public static BookUnit GetBookUnit(SpreadsheetDocument document)
        {
            return new BookUnit(document.WorkbookPart);
        }
    }
}
