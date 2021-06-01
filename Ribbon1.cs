using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelAddIn1
{
    public partial class Ribbon1
    {
        public int changeCount = 0;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private async void button1_Click(object sender, RibbonControlEventArgs e)
        {
            var activeWorksheet = ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
            
            var columnWidth = activeWorksheet.Columns[8].ColumnWidth;
            activeWorksheet.Columns[11].ColumnWidth = columnWidth;

            for (int i = 2; i < activeWorksheet.UsedRange.Rows.Count; i++)
            {
                // get source column value
                var Hvalue = activeWorksheet.Cells[i, 8].VALUE;
                //get result colum value
                var Kvalue = activeWorksheet.Cells[i, 11].VALUE;
                // get element column
                var Gvalue = activeWorksheet.Cells[i, 7].VALUE;

                //blow away result column value if any
                if (Kvalue != null)
                {
                    activeWorksheet.Cells[i, 11].VALUE = "";
                }
                if (Hvalue != null)
                {
                    //set formatting
                    activeWorksheet.Cells[i, 11].WrapText = true;
                    activeWorksheet.Cells[i, 11].Font.Name = "Arial";
                    activeWorksheet.Cells[i, 11].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                    activeWorksheet.Cells[i, 11].Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                    activeWorksheet.Cells[i, 11].Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
                    activeWorksheet.Cells[i, 11].Font.Size = 10;
                    activeWorksheet.Cells[i, 11].RowHeight = activeWorksheet.Cells[i, 8].RowHeight;
                    
                    //run cleanup operations
                    var res = await CleanUp(activeWorksheet.Cells[i, 8].VALUE);

                    
                    
                    // check if its result has any changes
                    //if (Hvalue == res) 
                    //    continue;

                    //if there is a difference, set the K col value
                    if(!string.IsNullOrEmpty(res))
                    {
                        activeWorksheet.Cells[i, 11].VALUE = res;
                    }

                    //review element count
                    if (Gvalue != null)
                    {
                        string elementResults = ElementCount(res, int.Parse(Gvalue));
                        if (elementResults.Length > 0)
                        {
                            activeWorksheet.Cells[i, 11].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            activeWorksheet.Cells[i, 11].VALUE += elementResults;
                        }
                    }
                    else activeWorksheet.Cells[i, 11].VALUE += " *NO ELEMENT COUNT";
                }
            }

            //popout change count
            if (changeCount > 0)
            {
                System.Windows.Forms.MessageBox.Show(changeCount + " Changes applied.");
            }
            else
                System.Windows.Forms.MessageBox.Show("No necessary changes were found.");
        }

        public async Task<string> CleanUp(string inputRow)
        {
            string outputRow = inputRow;
            outputRow = NumberPresentation(outputRow);
            outputRow = CleanWording(outputRow);
            outputRow = Punctuation(outputRow);

            return outputRow;
        }

        public string ElementCount(string input, int elementCount)
        {
            string output = "";
            var foundError = false;
            var highestElementInt = 0;
            
            //valid element counts
            var elementList = new List<string> { " 1\\) ", " 2\\) ", " 3\\) ", " 4\\) ", " 5\\) ", " 6\\) ", " 7\\) ", " 8\\) ", " 9\\) " };
            
            //check valid element counts
            foreach(var item in elementList)
            {
                var stripped = item.Replace(@"\", "").Replace(@"\", "");
                if (input.Contains(stripped))
                {
                    var occuranceCount = Regex.Matches(input, @item).Count;
                    if (occuranceCount > 1)
                    {
                        output += " *DUPLICATE ELEMENT NUMBERING. " + item + " occurs " + occuranceCount + " times.";
                        foundError = true;
                    }
                    
                    var elementInt = int.Parse(stripped.Trim().Replace(")", ""));
                    if (highestElementInt < elementInt) 
                        highestElementInt = elementInt;
                }
            }

            //check if the element count matches number of elements expected
            if (elementCount != highestElementInt)
            {
                if (highestElementInt > elementCount)
                    output += " *" + highestElementInt + " IS GREATER THAN EXPECTED COUNT " + elementCount;
                if (highestElementInt < elementCount)
                    output += " *" + highestElementInt + " IS LESS THAN EXPECTED COUNT " + elementCount;
            }

            //create list of bad elements
            var badElementList = new List<string>();
            for (int i = 10; i < 200; i++)
            {
                badElementList.Add(" " + i + ") ");
            }

            //check for bad elements
            foreach (var item in badElementList)
            {
                if (foundError) continue;
                if (input.Contains(item))
                {
                    output += " *BAD ELEMENT"+ item + "is > 9)";
                    foundError = true;
                }
            }

            return output;
        }

        public string Punctuation(string input)
        {
            var output = input;
            output = output.Trim(); 
            if (!output.EndsWith(".")) output += ".";
            output.Replace(",.", ".");
            output.Replace(".,", ".");
            output.Replace("  ", " ");
            output.Replace(" , ", ", ");
            output.Replace(" .", ".");
            output.Replace("..", ".");
            output.Replace(".)", ")");
            output.Replace(" ) ", ") ");
            output.Replace(" ( ", " (");
            output.Replace(" ; ", "; ");
            output = Regex.Replace(output, @"\r\n?|\n", "");

            if (output != input) changeCount++;
            return output;
        }

        public string NumberPresentation(string input)
        {
            var output = input;
            output.Replace("three hundred and sixty five (365)", "365");
            output.Replace("three hundred sixty-five (365)", "365");
            output.Replace("three hundred and sixty-five (365)", "365");
            output.Replace("three hundred and sixty- five (365)", "365");
            output.Replace("three hundred sixty five (365)", "365");
            output.Replace("three-hundred-sixty-five (365)", "365");
            output.Replace("365days", "365 days");
            output.Replace("[365]", "365");
            output.Replace("3 year", "three year");
            output.Replace("fifteen minute", "15 minute");
            output.Replace("fifty year", "50 year");
            output.Replace("thirty minute", "30 minute");
            output.Replace("twenty-four hour", "24 hour");
            output.Replace("twenty four hour", "24 hour");
            output.Replace("forty eight hour", "48 hour");
            output.Replace("forty-eight hour", "48 hour");
            output.Replace("seventy-two hour", "72 hour");
            output.Replace("seventy two hour", "72 hour");
            output.Replace("fourteen character", "14 character");
            output.Replace("7 day", "seven day");
            output.Replace("thirty day", "30 day");
            output.Replace("forty-five day", "45 day");
            output.Replace("forty five day", "45 day");
            output.Replace("sixty day", "60 day");
            output.Replace("ninety day", "90 day");
            output.Replace("3-year", "three year");
            output.Replace("fifteen minute", "15 minute");
            output.Replace("fifty year", "50 year");
            output.Replace("thirty minute", "30 minute");
            output.Replace("twenty four hour", "24 hour");
            output.Replace("forty eight hour", "48 hour");
            output.Replace("seventy two hour", "72 hour");
            output.Replace("fourteen character", "14 character");
            output.Replace("7 day", "seven day");
            output.Replace("thirty day", "30 day");
            output.Replace("forty five day", "45 day");
            output.Replace("sixty day", "60 day");
            output.Replace("ninety day", "90 day");

            if (output != input) changeCount++;
            return output;
        }

        public string CleanWording(string input)
        {
            var output = input;

            output.Replace("third party report", "third-party report");
            output.Replace("third party provider", "third-party provider");
            output.Replace("third party personnel", "third-party personnel");
            output.Replace("third party user", "third-party user");
            output.Replace("third party service", "third-party support");
            output.Replace("third party service", "third-party service");
            output.Replace("third party system", "third-party system");
            output.Replace("third party contact", "third-party contact");
            output.Replace("high risk locations", "high-risk locations");
            output.Replace("decision making roles", "decision-making roles");
            output.Replace("<p>", "");
            output.Replace("</p>", "");
            output.Replace("program(s) is (are)", "programs are");
            output.Replace("the organizations", "the organization's");
            output.Replace("third-parties", "third parties");
            output.Replace("counter-intelligence", "counterintelligence");
            output.Replace("personally-owned", "personally owned");
            output.Replace("up-to-date", "up to date");
            output.Replace("rol and", "role and");
            output.Replace("black list", "blacklist");
            output.Replace("internet", "Internet");
            output.Replace("rol, and", "role, and");
            output.Replace("controle ", "control ");
            output.Replace(" a updated", " an updated");
            output.Replace("activies", "activities");
            output.Replace("senor member", "senior member");
            output.Replace("endored", "endorsed");
            output.Replace("hard-drives", "hard drives");
            output.Replace("Group, shared or generic", "Group, shared, or generic");
            output.Replace("commonly-used", "commonly used");
            output.Replace("cryptographically-protected", "cryptographically protected");
            output.Replace("Visitor and third-party support access is recorded", "Visitor and third-party support access are recorded");

            if (output != input) changeCount++;
            return output;
        }


    }
}
