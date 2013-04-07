using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Packaging;
using System.Xml.Linq;
using System.IO;
using System.Xml;
using System.Data;

namespace Excel2Xml
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Cell> parsedCells = new List<Cell>();

            List<Row> parsedRows = new List<Row>();
            //string fileName = @"C:\Git\Projects\ICDMS\aaa.xlsx";
            string fileName = Path.Combine(Environment.CurrentDirectory, "aaa.xlsx");
            Package xlsxPackage = Package.Open(fileName, FileMode.Open, FileAccess.ReadWrite);
            try
            {
                PackagePartCollection allParts = xlsxPackage.GetParts();

                PackagePart sharedStringsPart = (from part in allParts
                                                 where part.ContentType.Equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")
                                                 select part).Single();

                XElement sharedStringsElement = XElement.Load(XmlReader.Create(sharedStringsPart.GetStream()));

                Dictionary<int, string> sharedStrings = new Dictionary<int, string>();
                ParseSharedStrings(sharedStringsElement, sharedStrings);

                XElement worksheetElement = GetWorksheet(1, allParts);

                IEnumerable<XElement> cells = from c in worksheetElement.Descendants(ExcelNamespaces.excelNamespace + "c")
                                              select c;

                foreach (XElement cell in cells)
                {
                    string cellPosition = cell.Attribute("r").Value;
                    int index = IndexOfNumber(cellPosition);
                    string column = cellPosition.Substring(0, index);
                    int row = Convert.ToInt32(cellPosition.Substring(index, cellPosition.Length - index));
                    if (cell.HasElements)
                    {
                        if (cell.Attribute("t") != null && cell.Attribute("t").Value == "s")
                        {
                            // Shared value
                            int valueIndex = Convert.ToInt32(cell.Descendants(ExcelNamespaces.excelNamespace + "v").Single().Value);
                            parsedCells.Add(new Cell(column, row, sharedStrings[valueIndex]));
                        }
                        else
                        {
                            string value = cell.Descendants(ExcelNamespaces.excelNamespace + "v").Single().Value;
                            parsedCells.Add(new Cell(column, row, value));
                        }
                    }
                    else
                    {
                        parsedCells.Add(new Cell(column, row, ""));
                    }
                }


                IEnumerable<XElement> rows = worksheetElement.Descendants(ExcelNamespaces.excelNamespace + "row");

                foreach (XElement row in rows)
                {
                    string rowID = row.Attribute("r").Value;

                    if (row.HasElements)
                    {
                        parsedRows.Add(new Row { RowID = rowID, Cells = parsedCells.Where(c => string.Compare(string.Format("{0}", c.Row), rowID, true) == 0 && !string.IsNullOrEmpty(c.Data)).ToArray() });
                    }
                }
            }
            finally
            {
                xlsxPackage.Close();
            }


            BuildXml bm = new BuildXml(parsedRows);
            bm.SetXml();
            //From here is additional code not covered in the posts, just to show it works
            //foreach (Cell cell in parsedCells)
            //{
            //    Console.WriteLine(cell);
            //    Console.ReadLine();
            //}
        }

        private static void ParseSharedStrings(XElement SharedStringsElement, Dictionary<int, string> sharedStrings)
        {
            IEnumerable<XElement> sharedStringsElements = from s in SharedStringsElement.Descendants(ExcelNamespaces.excelNamespace + "t")
                                                          select s;

            int Counter = 0;
            foreach (XElement sharedString in sharedStringsElements)
            {
                sharedStrings.Add(Counter, sharedString.Value);
                Counter++;
            }
        }

        private static XElement GetWorksheet(int worksheetID, PackagePartCollection allParts)
        {
            PackagePart worksheetPart = (from part in allParts
                                         where part.Uri.OriginalString.Equals(String.Format("/xl/worksheets/sheet{0}.xml", worksheetID))
                                         select part).Single();

            return XElement.Load(XmlReader.Create(worksheetPart.GetStream()));
        }

        private static int IndexOfNumber(string value)
        {
            for (int counter = 0; counter < value.Length; counter++)
            {
                if (char.IsNumber(value[counter]))
                {
                    return counter;
                }
            }

            return 0;
        }   
    }

    internal static class ExcelNamespaces
    {
        internal static XNamespace excelNamespace = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        internal static XNamespace excelRelationshipsNamepace = XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    }

    public class Cell
    {
        public Cell(string column, int row, string data)
        {
            this.Column = column;
            this.Row = row;
            this.Data = data;
        }

        public override string ToString()
        {
            return string.Format("{0}:{1} - {2}", Row, Column, Data);
        }

        public string Column { get; set; }
        public int Row { get; set; }
        public string Data { get; set; }
    }


    public class Row
    {
        public string RowID { get; set; }

        public Cell[] Cells { get; set; }
    }
}
