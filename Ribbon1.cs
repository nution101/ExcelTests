using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelAddIn1
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(
            //    this.button1_Click);
            


        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //Globals.ThisAddIn.Application. = true;
            //actionsPane2.Hide();
            //actionsPane1.Show();
            //toggleButton1.Checked = false;

            Microsoft.Office.Interop.Excel.Worksheet activeWorksheet = ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
            var rows = activeWorksheet.Rows;
            Console.WriteLine(rows.Count);
            //foreach (var row in rows.Value)
            //{
            //    var i = row.ToString();
            //    Console.WriteLine(i);
            //}


            DateTime dt = DateTime.Now;
            var rng = Globals.ThisAddIn.Application.get_Range("D2");
            object value = rng.Value2;

            if (value != null)
            {
                if (value is double)
                {
                    dt = DateTime.FromOADate((double)value);
                }
                else
                {
                    DateTime.TryParse((string)value, out dt);
                }
            }

            rng.Value = dt;

            //System.Windows.Forms.MessageBox.Show(dt.ToString());

            //firstRow.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
            //Microsoft.Office.Interop.Excel.Range newFirstRow = activeWorksheet.get_Range("A1");
            //newFirstRow.Value2 = "This text was added by using code";
        }

        
    }
}
