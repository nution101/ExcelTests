using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn1
{
    public partial class Ribbon1
    {
        public int changeCount = 0;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(
            //    this.button1_Click);
            


        }

        private async void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //Globals.ThisAddIn.Application. = true;
            //actionsPane2.Hide();
            //actionsPane1.Show();
            //toggleButton1.Checked = false;

            var activeWorksheet = ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
            //var rows = activeWorksheet.Rows;
            //Console.WriteLine(rows.Count);
            //foreach (var row in rows.Value)
            //{
            //    var i = row.ToString();
            //    Console.WriteLine(i);
            //}


            //DateTime dt = DateTime.Now;
            //var rng = Globals.ThisAddIn.Application.get_Range("D2");
            //object value = rng.Value2;

            //if (value != null)
            //{
            //    if (value is double)
            //    {
            //        dt = DateTime.FromOADate((double)value);
            //    }
            //    else
            //    {
            //        DateTime.TryParse((string)value, out dt);
            //    }
            //}

            //rng.Value = dt;

            //var colH = activeWorksheet.Columns[8];//Range of Column H
            //int totalColumns = activeWorksheet.UsedRange.Columns.Count;
            //int totalRows = activeWorksheet.UsedRange.Rows.Count;

            var columnWidth = activeWorksheet.Columns[8].ColumnWidth;
            activeWorksheet.Columns[11].ColumnWidth = columnWidth;

            for (int i = 2; i < activeWorksheet.UsedRange.Rows.Count; i++)
            {
                var Hvalue = activeWorksheet.Cells[i, 8].VALUE;
                if(Hvalue != null)
                {
                    //var rowJobj = activeWorksheet.Cells[i, 8];
                    //var rowKobj = activeWorksheet.Cells[i, 11];

                    activeWorksheet.Cells[i, 11].WrapText = true;
                    activeWorksheet.Cells[i, 11].Font.Name = "Arial";
                    activeWorksheet.Cells[i, 11].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                    activeWorksheet.Cells[i, 11].Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                    activeWorksheet.Cells[i, 11].Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
                    activeWorksheet.Cells[i, 11].Font.Size = 10;

                    activeWorksheet.Cells[i, 11].RowHeight = activeWorksheet.Cells[i, 8].RowHeight;
                    //activeWorksheet.Cells[i, 11].RowWidth = activeWorksheet.Cells[i, 8].RowWidth;
                    

                    var res = await CleanUp(activeWorksheet.Cells[i, 8].VALUE);

                    var originalText = activeWorksheet.Cells[i, 8].VALUE;
                    if (originalText == res) 
                        continue;

                    if(!string.IsNullOrEmpty(res))
                    {
                        activeWorksheet.Cells[i, 11].VALUE = res;
                    }
                    else
                    {
                        var test = 1;
                    }
                }
            }
            if (changeCount > 0)
            {
                System.Windows.Forms.MessageBox.Show(changeCount + " Changes applied");
            }

            //System.Windows.Forms.MessageBox.Show(dt.ToString());
            //firstRow.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
            //Microsoft.Office.Interop.Excel.Range newFirstRow = activeWorksheet.get_Range("A1");
            //newFirstRow.Value2 = "This text was added by using code";
        }

        public async Task<string> CleanUp(string inputRow)
        {
            string outputRow = inputRow;
            outputRow = Punctuation(outputRow);

            return outputRow;
        }

        public string Punctuation(string input)
        {
            var output = input;
            
            if (!output.EndsWith(".")) output += ".";
            output.Replace(",.", ".");
            output.Replace(".,", ".");
            output.Replace("  ", " ");
            output.Replace(" , ", ", ");
            output.Replace(" .", ".");
            output.Replace("..", ".");
            output.Replace(" ) ", ") ");
            output.Replace(" ( ", " (");
            output.Replace(" ; ", "; ");
            if (output != input) changeCount++;
            return output;
        }


    }
}
