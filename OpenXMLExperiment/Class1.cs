using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;

namespace OpenXMLExperiment
{
    public class Invoice
    {
        public string CnsldNo { get; set; }
        public string InvoiceNo { get; set; }
        public DateTime InvoiceDate { get; set; }
        public decimal Amount { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            string connectionString = "Server=DESKTOP-GCE98BC;Database=OpenXmlExcel;User Id=sa;Password=nopass;";

            // Fetch the data from the database
            List<Invoice> invoices = FetchDataFromDatabase(connectionString);

            // Manipulate the data
            List<Invoice> manipulatedInvoices = ManipulateData(invoices);

            // Write the data to an Excel file
            WriteExcelFileusigMemoryStream("new_excel_file2.xlsx", manipulatedInvoices);

            Console.WriteLine("OK");
            Console.ReadLine();

            //var data = new List<(string CnsldNo, string InvoiceNo, string InvoiceDate, int InvoiceAmount, int PaidAmount)>
            //{
            //    ("AAA111", "INV1", "12/11/2023", 100, 50),
            //    ("AAA111", "INV2", "12/12/2023", 200, 70),
            //    ("AAA111", "INV3", "12/13/2023", 300, 100),
            //    ("AAA112", "INV6", "12/11/2023", 100, 50),
            //    ("AAA112", "INV7", "12/12/2023", 200, 70),
            //    ("AAA113", "INV8", "12/11/2023", 100, 50),
            //    ("AAA113", "INV9", "12/12/2023", 200, 70)
            //};

            //using (SpreadsheetDocument document = SpreadsheetDocument.Create("Sample.xlsx", SpreadsheetDocumentType.Workbook))
            //{
            //    WorkbookPart workbookPart = document.AddWorkbookPart();
            //    workbookPart.Workbook = new Workbook();

            //    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            //    worksheetPart.Worksheet = new Worksheet(new SheetData());

            //    Sheets sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());

            //    Sheet sheet = new Sheet() { Id = document.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet 1" };
            //    sheets.Append(sheet);

            //    Worksheet worksheet = worksheetPart.Worksheet;
            //    SheetData sheetData = worksheet.GetFirstChild<SheetData>();

            //    // Create rows for each data item
            //    for (int i = 0; i < data.Count; i++)
            //    {
            //        var item = data[i];

            //        // Create a row
            //        Row row = new Row();

            //        // Create cells for each column
            //        Cell cnsldNoCell = new Cell() { DataType = CellValues.String, CellValue = new CellValue(item.CnsldNo) };
            //        Cell invoiceNoCell = new Cell() { DataType = CellValues.String, CellValue = new CellValue(item.InvoiceNo) };
            //        Cell invoiceDateCell = new Cell() { DataType = CellValues.String, CellValue = new CellValue(item.InvoiceDate) };
            //        Cell invoiceAmountCell = new Cell() { DataType = CellValues.Number, CellValue = new CellValue(item.InvoiceAmount.ToString()) };
            //        Cell paidAmountCell = new Cell() { DataType = CellValues.Number, CellValue = new CellValue(item.PaidAmount.ToString()) };

            //        // Create a cell with a formula for Balance Amount
            //        Cell balanceAmountCell = new Cell();
            //        //balanceAmountCell.CellReference = $"F{i + 2}";
            //        balanceAmountCell.DataType = CellValues.Number;
            //        balanceAmountCell.CellFormula = new CellFormula($"D{i + 1}-E{i + 1}"); // Balance Amount = Invoice Amount - Paid Amount
            //        balanceAmountCell.CellValue = new CellValue();

            //        // Add the cells to the row
            //        row.Append(cnsldNoCell, invoiceNoCell, invoiceDateCell, invoiceAmountCell, paidAmountCell, balanceAmountCell);

            //        // Add the row to the sheet data
            //        sheetData.AppendChild(row);
            //    }

            //    workbookPart.Workbook.Save();
            //}
        }

        public static List<Invoice> FetchDataFromDatabase(string connectionString)
        {
            var invoices = new List<Invoice>();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("SELECT * FROM dbo.Invoices", connection))
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        invoices.Add(new Invoice
                        {
                            CnsldNo = reader.GetValue(1) != null ? reader.GetValue(1).ToString() : string.Empty,
                            InvoiceNo = reader.GetString(2),
                            InvoiceDate = reader.GetDateTime(3),
                            Amount = reader.GetDecimal(4)
                        });
                    }
                }
            }

            return invoices;
        }

        public static List<Invoice> ManipulateData(List<Invoice> invoices)
        {
            // Manipulate the data according to your needs
            // For example, remove duplicate "Cnsld No" values
            var filteredCnsldInvoices = new List<Invoice>();

            var cnsldInvoices = invoices.Where(x => x.CnsldNo.Trim() != "").ToList();
            
            var unCnsldinvoices = invoices.Where(x => x.CnsldNo == "" || x.CnsldNo == null).ToList();

            if(cnsldInvoices!= null)
            {
                var cnsldInvoiceGroup = cnsldInvoices.GroupBy(x => x.CnsldNo.Trim()).ToList();
                foreach (var item in cnsldInvoiceGroup)
                {
                    var seletedCnsldInvoice = cnsldInvoices.Where(x => x.CnsldNo.Trim() == item.Key).ToList();
                    for (int j = 0; j < seletedCnsldInvoice.Count; j++)
                    {
                        if (j == 0)
                        {
                            filteredCnsldInvoices.Add(seletedCnsldInvoice[j]);
                        }
                        else
                        {
                            seletedCnsldInvoice[j].CnsldNo = "";
                            filteredCnsldInvoices.Add(seletedCnsldInvoice[j]);
                        }
                    }
                }
            }

            if(unCnsldinvoices != null)
            {
                filteredCnsldInvoices.AddRange(unCnsldinvoices);
            }


            return filteredCnsldInvoices;
        }


        public static void WriteExcelFileusigMemoryStream(string fileName, List<Invoice> invoices)
        {
            using (MemoryStream mem = new MemoryStream())
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Create(mem, SpreadsheetDocumentType.Workbook))
                {
                    // Create a new fill object with the desired color.
                    Fill fill = new Fill()
                    {
                        PatternFill = new PatternFill()
                        {
                            PatternType = PatternValues.Solid,
                            BackgroundColor = new BackgroundColor()
                            {
                                Rgb = new HexBinaryValue() { Value = "FFFF0000" } // Red color
                            }
                        }
                    };


                    WorkbookPart workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());

                    // Add styles to the workbook
                    WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                    workbookStylesPart.Stylesheet = GenerateStyleSheet();
                    //workbookStylesPart.Stylesheet.Fills = new Fills();
                    //workbookStylesPart.Stylesheet.Fills.Append(fill);
                    workbookStylesPart.Stylesheet.Save();

                    // Create custom widths for columns
                    Columns lstColumns = worksheetPart.Worksheet.GetFirstChild<Columns>();
                    Boolean needToInsertColumns = false;
                    if (lstColumns == null)
                    {
                        lstColumns = new Columns();
                        needToInsertColumns = true;
                    }
                    // Min = 1, Max = 1 ==> Apply this to column 1 (A)
                    // Min = 2, Max = 2 ==> Apply this to column 2 (B)
                    // Width = 25 ==> Set the width to 25
                    // CustomWidth = true ==> Tell Excel to use the custom width
                    lstColumns.Append(new Column() { Min = 1, Max = 1, Width = 25, CustomWidth = true });
                    lstColumns.Append(new Column() { Min = 2, Max = 2, Width = 20, CustomWidth = true });
                    lstColumns.Append(new Column() { Min = 3, Max = 3, Width = 20, CustomWidth = true });
                    lstColumns.Append(new Column() { Min = 4, Max = 4, Width = 20, CustomWidth = true });
                    // Only insert the columns if we had to create a new columns element
                    if (needToInsertColumns)
                        worksheetPart.Worksheet.InsertAt(lstColumns, 0);


                    // Create a new cell format that references the fill.
                    //CellFormat cellFormat = new CellFormat() { FillId = 0 }; // 0 is the index of our fill in the fills collection.

                    // Add the cell format to the cell formats collection of the workbook styles part.
                    //workbookStylesPart.Stylesheet.CellFormats = new CellFormats();
                    //workbookStylesPart.Stylesheet.CellFormats.Append(cellFormat);


                    Sheets sheets = document.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                    Sheet sheet = new Sheet() { Id = document.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Invoices" };
                    sheets.Append(sheet);

                    SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                    // Add the headers
                    Row row = new Row();
                    row.Append(
                        new Cell() { DataType = CellValues.String, CellValue = new CellValue("Cnsld No"), StyleIndex = 4 },
                        new Cell() { DataType = CellValues.String, CellValue = new CellValue("Invoice no"), StyleIndex = 4 },
                        new Cell() { DataType = CellValues.String, CellValue = new CellValue("Invoice date"), StyleIndex = 4 },
                        new Cell() { DataType = CellValues.String, CellValue = new CellValue("Amount"), StyleIndex = 4 }
                    );
                    sheetData.AppendChild(row);

                    // Add the data
                    foreach (Invoice invoice in invoices)
                    {
                        row = new Row();
                        row.Append(
                            new Cell() { DataType = CellValues.String, CellValue = new CellValue(invoice.CnsldNo) },
                            new Cell() { DataType = CellValues.String, CellValue = new CellValue(invoice.InvoiceNo) },
                            new Cell() { DataType = CellValues.String, CellValue = new CellValue(invoice.InvoiceDate.ToString("MM-dd-yyyy")) },
                            new Cell() { DataType = CellValues.Number, CellValue = new CellValue(invoice.Amount.ToString("#.####")) }
                        );
                        sheetData.AppendChild(row);
                    }
                    document.Close();
                }

                FileStream fileStream = new FileStream(AppDomain.CurrentDomain.BaseDirectory + $"{fileName}.xlsx", FileMode.Create, FileAccess.Write);
                mem.WriteTo(fileStream);
                fileStream.Close();
                mem.Close();
            }
        }

        public static void WriteExcelFile(string path, List<Invoice> invoices)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
            {
                // Create a new fill object with the desired color.
                Fill fill = new Fill()
                {
                    PatternFill = new PatternFill()
                    {
                        PatternType = PatternValues.Solid,
                        ForegroundColor = new ForegroundColor()
                        {
                            Rgb = new HexBinaryValue() { Value = "FFFF0000" } // Red color
                        }
                    }
                };


                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());


                // Create Columns object
                Columns columns = new Columns();

                // Create Column objects with custom width and append them to the Columns object
                columns.Append(new Column() { Min = 1, Max = 1, Width = 50, CustomWidth = true });
                columns.Append(new Column() { Min = 2, Max = 2, Width = 50, CustomWidth = true });
                columns.Append(new Column() { Min = 3, Max = 3, Width = 50, CustomWidth = true });
                columns.Append(new Column() { Min = 4, Max = 4, Width = 50, CustomWidth = true });

                // Append the Columns object to the Worksheet
                worksheetPart.Worksheet.Append(columns);

                // Add styles to the workbook
                WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                workbookStylesPart.Stylesheet = new Stylesheet();
                workbookStylesPart.Stylesheet.Fills = new Fills();
                workbookStylesPart.Stylesheet.Fills.Append(fill);

                // Create a new cell format that references the fill.
                CellFormat cellFormat = new CellFormat() { FillId = 0 }; // 0 is the index of our fill in the fills collection.

                // Add the cell format to the cell formats collection of the workbook styles part.
                workbookStylesPart.Stylesheet.CellFormats = new CellFormats();
                workbookStylesPart.Stylesheet.CellFormats.Append(cellFormat);


                Sheets sheets = document.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                Sheet sheet = new Sheet() { Id = document.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Invoices" };
                sheets.Append(sheet);

                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                // Add the headers
                Row row = new Row();
                row.Append(
                    new Cell() { DataType = CellValues.String, CellValue = new CellValue("Cnsld No"), StyleIndex=4 },
                    new Cell() { DataType = CellValues.String, CellValue = new CellValue("Invoice no"), StyleIndex = 4 },
                    new Cell() { DataType = CellValues.String, CellValue = new CellValue("Invoice date"), StyleIndex = 4 },
                    new Cell() { DataType = CellValues.String, CellValue = new CellValue("Amount"), StyleIndex = 4 }
                );
                sheetData.AppendChild(row);

                // Add the data
                foreach (Invoice invoice in invoices)
                {
                    row = new Row();
                    row.Append(
                        new Cell() { DataType = CellValues.String, CellValue = new CellValue(invoice.CnsldNo) },
                        new Cell() { DataType = CellValues.String, CellValue = new CellValue(invoice.InvoiceNo) },
                        new Cell() { DataType = CellValues.String, CellValue = new CellValue(invoice.InvoiceDate.ToString("MM-dd-yyyy")) },
                        new Cell() { DataType = CellValues.Number, CellValue = new CellValue(invoice.Amount.ToString("#.####")) }
                    );
                    sheetData.AppendChild(row);
                    
                }
                workbookPart.Workbook.Save();
            }
        }


        public static Stylesheet GenerateStyleSheet()
        {
            return new Stylesheet(
            new Fonts(
            new Font(new FontSize() { Val = 11 }, new Color() { Rgb = new HexBinaryValue() { Value = "000000" } }, new FontName() { Val = "Calibri" }),// Index 0 - The default font.
            new Font(new Bold(), new FontSize() { Val = 11 }, new Color() { Rgb = new HexBinaryValue() { Value = "000000" } }, new FontName() { Val = "Calibri" }),  // Index 1 - The bold font.
            new Font(new Italic(), new FontSize() { Val = 11 }, new Color() { Rgb = new HexBinaryValue() { Value = "000000" } }, new FontName() { Val = "Calibri" }),  // Index 2 - The Italic font.
            new Font(new FontSize() { Val = 18 }, new Color() { Rgb = new HexBinaryValue() { Value = "000000" } }, new FontName() { Val = "Calibri" }),  // Index 3 - The Times Roman font. with 16 size
            new Font(new Bold(), new FontSize() { Val = 18 }, new Color() { Rgb = new HexBinaryValue() { Value = "000000" } }, new FontName() { Val = "Calibri" }),  // Index 4 - The Times Roman font. with 16 size
            new Font(new Bold(), new FontSize() { Val = 11 }, new Color() { Rgb = new HexBinaryValue() { Value = "FFFFFF" } }, new FontName() { Val = "Calibri" }),  // Index 5 - The bold font.
            new Font(new FontSize() { Val = 11 }, new Color() { Rgb = new HexBinaryValue() { Value = "FFFFFF" } }, new FontName() { Val = "Calibri" })  // Index 6 - The normal font with white font color.

            ),
            new Fills(
            new Fill( // Index 0 - The default fill.
            new DocumentFormat.OpenXml.Spreadsheet.PatternFill() { PatternType = PatternValues.None }),
            new DocumentFormat.OpenXml.Spreadsheet.Fill( // Index 1 - The default fill of gray 125 (required)
            new DocumentFormat.OpenXml.Spreadsheet.PatternFill() { PatternType = PatternValues.Gray125 }),
            new DocumentFormat.OpenXml.Spreadsheet.Fill( // Index 2 - The yellow fill.
            new DocumentFormat.OpenXml.Spreadsheet.PatternFill(
            new DocumentFormat.OpenXml.Spreadsheet.ForegroundColor() { Rgb = new HexBinaryValue() { Value = "1f497d" } }
            )
            { PatternType = PatternValues.Solid }),
            new DocumentFormat.OpenXml.Spreadsheet.Fill( // Index 3 - The Blue fill.
            new DocumentFormat.OpenXml.Spreadsheet.PatternFill(
            new DocumentFormat.OpenXml.Spreadsheet.ForegroundColor() { Rgb = new HexBinaryValue() { Value = "8EA9DB" } }
            )
            { PatternType = PatternValues.Solid })
            ),
            new Borders(
            new Border( // Index 0 - The default border.
            new DocumentFormat.OpenXml.Spreadsheet.LeftBorder(),
            new DocumentFormat.OpenXml.Spreadsheet.RightBorder(),
            new DocumentFormat.OpenXml.Spreadsheet.TopBorder(),
            new DocumentFormat.OpenXml.Spreadsheet.BottomBorder(),
            new DiagonalBorder()),
            new Border( // Index 1 - Applies a Left, Right, Top, Bottom border to a cell
            new DocumentFormat.OpenXml.Spreadsheet.LeftBorder(
            new Color() { Auto = true }
            )
            { Style = BorderStyleValues.Thin },
            new DocumentFormat.OpenXml.Spreadsheet.RightBorder(
            new Color() { Auto = true }
            )
            { Style = BorderStyleValues.Thin },
            new DocumentFormat.OpenXml.Spreadsheet.TopBorder(
            new Color() { Auto = true }
            )
            { Style = BorderStyleValues.Thin },
            new DocumentFormat.OpenXml.Spreadsheet.BottomBorder(
            new Color() { Auto = true }
            )
            { Style = BorderStyleValues.Thin },
            new DiagonalBorder()),
                   new Border( // Index 1 - Applies a Left, Right, Top, Bottom border to a cell
            new DocumentFormat.OpenXml.Spreadsheet.LeftBorder(
            new Color() { Auto = true }
            )
            { Style = BorderStyleValues.None },
            new DocumentFormat.OpenXml.Spreadsheet.RightBorder(
            new Color() { Auto = true }
            )
            { Style = BorderStyleValues.None },
            new DocumentFormat.OpenXml.Spreadsheet.TopBorder(
            new Color() { Auto = true }
            )
            { Style = BorderStyleValues.None },
            new DocumentFormat.OpenXml.Spreadsheet.BottomBorder(
            new Color() { Rgb = new HexBinaryValue() { Value = "FFA500" } }
            )
            { Style = BorderStyleValues.Thin },
            new DiagonalBorder())
            ),
            new CellFormats(
            new CellFormat() { FontId = 0, FillId = 0, BorderId = 0 }, // Index 0 - The default cell style. If a cell does not have a style index applied it will use this style combination instead
            new CellFormat() { FontId = 1, FillId = 0, BorderId = 0, ApplyFont = true }, // Index 1 - Bold
            new CellFormat() { FontId = 2, FillId = 0, BorderId = 0, ApplyFont = true }, // Index 2 - Italic
            new CellFormat() { FontId = 3, FillId = 0, BorderId = 0, ApplyFont = true }, // Index 3 - Times Roman
            new CellFormat() { FontId = 6, FillId = 2, BorderId = 0, ApplyFill = true }, // Index 4 - Yellow Fill
            new CellFormat( // Index 5 - Alignment
            new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
            )
            { FontId = 0, FillId = 0, BorderId = 0, ApplyAlignment = true },
            new CellFormat() { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }, // Index 6 - Border
             new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }) // Index 7 - Alignment
             { FontId = 1, FillId = 0, BorderId = 0, ApplyAlignment = true },

             new CellFormat() { FontId = 4, FillId = 0, BorderId = 0, ApplyFont = true }, // Index 8 - Times Roman
             new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }) { FontId = 0, FillId = 0, BorderId = 2, ApplyFont = true }, // Index 9 - Bottom Border with Color 70AD47
             new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }) // Index 10 - Alignment
             { FontId = 5, FillId = 3, BorderId = 0, ApplyAlignment = true }


             )
            ); // return
        }
    }
}
