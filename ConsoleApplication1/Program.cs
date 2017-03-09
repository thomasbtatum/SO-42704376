using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            Create();
        }


        public static void Create()
        {
            string appPath = System.IO.Path.GetDirectoryName(System.IO.Path.GetDirectoryName(System.IO.Directory.GetCurrentDirectory()));

            string templateFile = appPath + @"\Templates\ChartExample.xlsx";
            string saveFile = appPath + @"\Documents\Generated2.xlsx";



            File.Copy(templateFile, saveFile, true);

            //open copied template.
            using (SpreadsheetDocument myWorkbook = SpreadsheetDocument.Open(saveFile, true))
            {
                //this is the workbook contains all the worksheets
                WorkbookPart workbookPart = myWorkbook.WorkbookPart;

                //we know that the first worksheet contains the data for the graph
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First(); //getting the first worksheet
                                                                                   //the shhet data contains the information we are looking to alter
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                int index = 2;//Row the data for the graph starts on
                              //var qry = from t in db.SEL_SE_DEATHS()
                FudgeData fudge = new FudgeData();

                List<FudgeItem> qry = fudge.Fudged();

                foreach (FudgeItem item in qry)
                {
                    int Year = item.EventYear;
                    int PSQ = item.PSQReviewable;
                    int death = item.Deaths;

                    Row contentRow = CreateContentRow(index, Year, PSQ, death);
                    index++;
                    //contentRow.RowIndex = (UInt32)index;
                    sheetData.AppendChild(contentRow);

                }

                //(<x:c r="A2" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:v>2014</x:v></x:c><x:c r="B2" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:v>21</x:v></x:c><x:c r="C2" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:v>4</x:v></x:c>)
                FixChartData(workbookPart, index);
                worksheetPart.Worksheet.Save();

                myWorkbook.Close();
                myWorkbook.Dispose();
            }

        }

        static string[] headerColumns = new string[] { "A", "B", "C" }; //the columns being accessed
        public static Row CreateContentRow(int index, int year, int pSQ, int death)
        {
            Row r = new Row();
            r.RowIndex = (UInt32)index;

            //skipping the text add function

            //we are createing a cell for each column (headerColumns),
            //for each cell we are adding a value.
            //we then append the value to the cell and append the cell to the row - wich is returned.
            for (int i = 0; i < headerColumns.Length; i++)
            {
                Cell c = new Cell();
                c.CellReference = headerColumns[i] + index;
                CellValue v = new CellValue();
                if (i == 0)
                {
                    v.Text = year.ToString();
                }
                else if (i == 1)
                {
                    v.Text = pSQ.ToString();
                }
                else if (i == 2)
                {
                    v.Text = death.ToString();
                }
                c.AppendChild(v);
                r.AppendChild(c);
            }
            return r;

        }

        //Method for when the datatype is text based
        public Cell CreateTextCell(string header, string text, int index)
        {
            //Create a new inline string cell.
            Cell c = new Cell();
            c.DataType = CellValues.InlineString;
            c.CellReference = header + index;
            //Add text to the text cell.
            InlineString inlineString = new InlineString();
            Text t = new Text();
            t.Text = text;
            inlineString.AppendChild(t);
            c.AppendChild(inlineString);
            return c;
        }

        //fix the chart Data Regions
        public static void FixChartData(WorkbookPart workbookPart, int totalCount)
        {

            var wsparts = workbookPart.WorksheetParts.ToArray();

            foreach (WorksheetPart wsp in wsparts)
            {
                if (wsp.DrawingsPart != null)
                {
                    ChartPart chartPart = wsp.DrawingsPart.ChartParts.First();
                    ////change the ranges to accomodate the newly inserted data.
                    foreach (DocumentFormat.OpenXml.Drawing.Charts.Formula formula in chartPart.ChartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Formula>())
                    {
                        if (formula.Text.Contains("$2"))
                        {
                            string s = formula.Text.Split('$')[1];
                            formula.Text += ":$" + s + "$" + totalCount;
                        }
                    }
                    chartPart.ChartSpace.Save();
                }
            }

            //ChartPart chartPart = workbookPart.ChartsheetParts.First().DrawingsPart.ChartParts.First();
            ////change the ranges to accomodate the newly inserted data.
            //foreach (DocumentFormat.OpenXml.Drawing.Charts.Formula formula in chartPart.ChartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Formula>())
            //{
            //    if (formula.Text.Contains("$2"))
            //    {
            //        string s = formula.Text.Split('$')[1];
            //        formula.Text += ":$" + s + "$" + totalCount;
            //    }
            //}
            //chartPart.ChartSpace.Save();
        }
    }
}
