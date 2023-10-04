using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Utility.Models;
using Document = DocumentFormat.OpenXml.Wordprocessing.Document;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;

namespace Utility.Pages
{
    public class FileExtractionModel : PageModel
    {
        private readonly IWebHostEnvironment _environment;

        public FileExtractionModel(IWebHostEnvironment environment)
        {
            _environment = environment;
        }
        public IActionResult OnGet()
        {
            return Page();
        }
        [BindProperty]
        public FileExtract FileExtract { get; set; } = default!;

        public class ExtractData
        {
            public string TestName { get; set; }
            public string TestResult { get; set; }
            public string Time { get; set; }
        }
        public async Task<IActionResult> OnPostAsync()
        {
            var files = FileExtract.files;

            if (files != null && files.Count > 0)
            {
                //To remove RTF tags
                string[] wordsToRemove = { @"\viewkind4", @"\uc1", @"\pard", @"\f0", @"\fs20", @"\cf0", @"\cf1", @"\cf2", @"\cf3", @"\cf4", @"\rtlch", @"\fcs1", 
                    @"\af43", @"\afs20", @"\ltrch", @"\fcs0", @"\f431", @"\kerning0",@"{",@"}", @"\insrsid3868228", @"\hich", @"\af43", @"\dbch", @"\af31505", @"\loch", @"\f43" };//@"\par",s
                List<ExtractData> extractData = new List<ExtractData>();
                foreach (var file in files)
                { 
                    if (file.Length > 0)
                    {

                        //Read and convert RTF content to plain text
                        string rtfContent;
                        using (var streamReader = new StreamReader(file.OpenReadStream()))
                        {
                            rtfContent = streamReader.ReadToEnd();
                        }
                        
                        string modifiedString = string.Join("", rtfContent.Split(wordsToRemove, StringSplitOptions.None));

                        //Converting as Lines
                        string[] lines = modifiedString.Split(new string[] { "\\par" }, StringSplitOptions.None);
                        int index = 0;
                        foreach(var line in lines)
                        {
                            lines[index] = line.Replace("\r", string.Empty).Replace("\n", string.Empty);
                            index++;
                        }
                                // Extract values
                                string testName = FindNextLine(lines, "The CAN speed");
                                string testResult = FindNextLine(lines, "the test results are");
                        string timeResult = FindNextLine(lines, "The total test time was measured ");

                        extractData.Add(new ExtractData { TestName = testName,TestResult= testResult, Time=  timeResult });
                    }
                }
                #region DataAppendToFile
                //WordZone
                // Create a new Word document
                MemoryStream memoryStream = new MemoryStream();
                using (WordprocessingDocument doc = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = doc.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());

                    // Create a table
                    Table table = new Table();
                    TableProperties tableProperties = new TableProperties(new TableBorders(
                        new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                        new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                        new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                        new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 }
                    ));
                    table.AppendChild(tableProperties);

                    // Add table headers
                    TableRow headerRow = new TableRow();
                    headerRow.Append(CreateHeaderCell("S.No"));
                    headerRow.Append(CreateHeaderCell("Test Name"));
                    headerRow.Append(CreateHeaderCell("Test Result"));
                    headerRow.Append(CreateHeaderCell("Test Time"));
                    table.AppendChild(headerRow);

                    // Add table data
                    for (int i = 0; i < extractData.Count; i++)
                    {
                        TableRow dataRow = new TableRow();
                        dataRow.Append(CreateCell((i + 1).ToString())); // S.No
                        dataRow.Append(CreateCell(extractData[i].TestName)); // Test Name
                        dataRow.Append(CreateCell(extractData[i].TestResult)); // Test Result
                        dataRow.Append(CreateCell(extractData[i].Time)); // Test Time
                        table.AppendChild(dataRow);
                    }

                    body.Append(table);
                }

                // Set the response headers for downloading the file
                Response.Headers.Add("Content-Disposition", "attachment; filename=ExtractResult.docx");
                return new FileStreamResult(new MemoryStream(memoryStream.ToArray()), "application/vnd.openxmlformats-officedocument.wordprocessingml.document");

                #endregion
                #region ExcelTemplate
                ////ExcelZone
                //// Create a new Excel workbook and sheet
                //IWorkbook workbook = new XSSFWorkbook();
                //ISheet sheet = workbook.CreateSheet("TestResults");

                //// Create a header row with bold formatting
                //ICellStyle headerStyle = workbook.CreateCellStyle();
                //IFont headerFont = workbook.CreateFont();
                //headerFont.Boldweight = (short)FontBoldWeight.Bold;
                //headerStyle.SetFont(headerFont);

                //IRow headerRow1 = sheet.CreateRow(0);
                //headerRow1.CreateCell(0).SetCellValue("S.No");
                //headerRow1.CreateCell(1).SetCellValue("Test Name");
                //headerRow1.CreateCell(2).SetCellValue("Test Result");
                //headerRow1.CreateCell(3).SetCellValue("Test Time");

                //foreach (var cell in headerRow1.Cells)
                //{
                //    cell.CellStyle = headerStyle;
                //}

                //// Populate the data
                //var data = extractData;
                //for (int i = 0; i < data.Count; i++)
                //{
                //    IRow dataRow = sheet.CreateRow(i + 1);
                //    dataRow.CreateCell(0).SetCellValue(i + 1);
                //    dataRow.CreateCell(1).SetCellValue(data[i].TestName);
                //    dataRow.CreateCell(2).SetCellValue(data[i].TestResult);
                //    dataRow.CreateCell(3).SetCellValue(data[i].Time);
                //}

                //// Create a memory stream and write the workbook to it
                //MemoryStream memoryStream2 = new MemoryStream();
                //workbook.Write(memoryStream2);

                //// Set the response headers for downloading the file
                //Response.Headers.Add("Content-Disposition", "attachment; filename=TestResults.xlsx");
                //return new FileStreamResult(new MemoryStream(memoryStream2.ToArray()), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

                #endregion

            }
            return RedirectToPage("/SuccessPage");
        }

        static string FindNextLine(string[] lines, string searchText)
        {
            bool found = false;

            foreach (string line in lines)
            {
                if (found)
                {
                    // Return the next line after the line with the provided text
                    return line;
                }

                if (line.Trim().Contains(searchText.Trim()))
                {
                    found = true;
                }
            }

            // If the provided text was not found or it's the last line, return null or an empty string
            return null; // or return string.Empty;
        }
        private TableCell CreateHeaderCell(string text)
        {
            TableCell cell = new TableCell();
            Paragraph paragraph = new Paragraph(new Run(new Text(text)));
            paragraph.ParagraphProperties = new ParagraphProperties(new Justification { Val = JustificationValues.Center });
            cell.Append(paragraph);
            return cell;
        }

        private TableCell CreateCell(string text)
        {
            TableCell cell = new TableCell();
            Paragraph paragraph = new Paragraph(new Run(new Text(text)));
            paragraph.ParagraphProperties = new ParagraphProperties(new Justification { Val = JustificationValues.Left });
            cell.Append(paragraph);
            return cell;
        }

    }
}
