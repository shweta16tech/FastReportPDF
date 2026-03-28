using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FastReport;
using FastReport.Export.PdfSimple;
using System;
using System.IO;
namespace TestReport
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // -------- PDF only ----- **** USING FAST REPORT *******
            //string reportPath = Path.Combine(Directory.GetCurrentDirectory(), @"Reports\FirstReport.frx");

            //Report reporting = Report.FromFile(reportPath);

            //// Parameters to the report 
            //reporting.SetParameterValue("Name", "John");
            //reporting.SetParameterValue("Email", "john123@gmail.com");
            //reporting.SetParameterValue("Message", "Hello everyone !!");


            //reporting.Prepare();

            //using var pdfExport = new FastReport.Export.PdfSimple.PDFSimpleExport();
            //using var reportStream = new MemoryStream();
            //pdfExport.Export(reporting, reportStream);
            //File.WriteAllBytes("HereisPdf.PDF", reportStream.ToArray());



            string reportPath = Path.Combine(Directory.GetCurrentDirectory(), @"Reports\FirstReport.frx");

            // Load report
            Report reporting = Report.FromFile(reportPath);

            // Set report parameters
            reporting.SetParameterValue("Name", "John");
            reporting.SetParameterValue("Email", "john123@gmail.com");
            reporting.SetParameterValue("Message", "Hello everyone !!");

            reporting.Prepare();

            // ---------------- PDF Export ----------------
            using (var pdfExport = new PDFSimpleExport())
            using (var reportStream = new MemoryStream())
            {
                pdfExport.Export(reporting, reportStream);
                File.WriteAllBytes("HereIsPdf.pdf", reportStream.ToArray());
                Console.WriteLine("PDF exported successfully.");
            }

            // ---------------- Excel Export ----------------
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Report");

                // Add headers
                ws.Cell(1, 1).Value = "Name";
                ws.Cell(1, 2).Value = "Email";
                ws.Cell(1, 3).Value = "Message";

                // Add values safely (handle nulls)
                ws.Cell(2, 1).Value = reporting.GetParameterValue("Name")?.ToString() ?? "";
                ws.Cell(2, 2).Value = reporting.GetParameterValue("Email")?.ToString() ?? "";
                ws.Cell(2, 3).Value = reporting.GetParameterValue("Message")?.ToString() ?? "";

                workbook.SaveAs("HereIsExcel.xlsx");
                Console.WriteLine("Excel exported successfully.");
            }

            // ---------------- Word Export ----------------
            using (var wordDoc = WordprocessingDocument.Create("HereIsWord.docx", WordprocessingDocumentType.Document))
            {
                var mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                // Create table
                Table table = new Table();

                // Table borders
                TableProperties tblProps = new TableProperties(
                    new TableBorders(
                        new TopBorder { Val = BorderValues.Single, Size = 6 },
                        new BottomBorder { Val = BorderValues.Single, Size = 6 },
                        new LeftBorder { Val = BorderValues.Single, Size = 6 },
                        new RightBorder { Val = BorderValues.Single, Size = 6 },
                        new InsideHorizontalBorder { Val = BorderValues.Single, Size = 6 },
                        new InsideVerticalBorder { Val = BorderValues.Single, Size = 6 }
                    )
                );
                table.AppendChild(tblProps);

                // Header row
                TableRow headerRow = new TableRow();
                headerRow.Append(
                    new TableCell(new Paragraph(new Run(new Text("Name")))),
                    new TableCell(new Paragraph(new Run(new Text("Email")))),
                    new TableCell(new Paragraph(new Run(new Text("Message"))))
                );
                table.AppendChild(headerRow);

                // Data row
                TableRow dataRow = new TableRow();
                dataRow.Append(
                    new TableCell(new Paragraph(new Run(new Text(reporting.GetParameterValue("Name")?.ToString() ?? "")))),
                    new TableCell(new Paragraph(new Run(new Text(reporting.GetParameterValue("Email")?.ToString() ?? "")))),
                    new TableCell(new Paragraph(new Run(new Text(reporting.GetParameterValue("Message")?.ToString() ?? ""))))
                );
                table.AppendChild(dataRow);

                body.AppendChild(table);
                Console.WriteLine("Word exported successfully.");
            }

            Console.WriteLine("All exports completed!");


        }
    }
}