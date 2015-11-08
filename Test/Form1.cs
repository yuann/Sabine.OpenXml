using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Sabine.OpenXml;

namespace Test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string fileName;
            using (var dlg = new OpenFileDialog())
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;
                fileName = dlg.FileName;
            }

            using (var document = SpreadsheetDocument.Open(fileName, false))
            {
                var bookUnit = SpreadsheetEx.GetBookUnit(document);

                var cellValues = bookUnit.SheetUnits
                    .Where(x => x.Name == "Sheet1")
                    .SelectMany(x => x.RowUnits)
                    .SelectMany(x => x.CellUnits)
                    .Select(x => x.GetString());

                Console.WriteLine(string.Join("\t", cellValues));

                var sheetUnit = bookUnit.SheetUnits.Single(x => x.Name == "Sheet1");

                var columns = sheetUnit.RowUnits
                    .Where(x => x.RowNumber < 3)
                    .SelectMany(x => x.CellUnits)
                    .ToLookup(x => x.ColumnNumber)
                    .Select(g => new Column()
                    {
                        Item1 = g.SingleOrDefault(x => x.RowNumber == 1).GetString(),
                        Item2 = g.SingleOrDefault(x => x.RowNumber == 2).GetString()
                    });
            }
        }

        class Column
        {
            public string Item1 { get; set; }
            public string Item2 { get; set; }
        }
    }
}
