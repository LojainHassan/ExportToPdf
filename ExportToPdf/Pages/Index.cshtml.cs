using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using ExportToPdf.Models;
using ExportToPdf.Services;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QColors = QuestPDF.Helpers.Colors;
using QuestPDF.Infrastructure;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace ExportToPdf.Pages
{
    public class IndexModel : PageModel
    {
        private readonly SearchService _searchService;

        public IndexModel(SearchService searchService)
        {
            _searchService = searchService;
        }

        public IEnumerable<SearchModel> SearchModels { get; set; }

        public void OnGet()
        {
            SearchModels = _searchService.GetSearchModels();
        }

        public IActionResult OnGetDownloadPdf()
        {
            var data = _searchService.GetSearchModels().ToList();
            var chunks = data.Chunk(20).ToList();

            var document = Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Margin(1, Unit.Centimetre);
                    page.Size(PageSizes.A4);
                    page.ContentFromRightToLeft();

                    // تصغير الخط العام
                    page.DefaultTextStyle(x => x.FontFamily("Times New Roman").FontSize(9));
                    // Header
                    page.Header()
                        .PaddingBottom(10)
                        .Text("تصدير السجلات المرجعية")
                        .SemiBold()
                        .AlignCenter()
                        .FontSize(12)
                        .FontColor(QColors.Black);

                    // Content
                    page.Content().Column(column =>
                    {
                        foreach (var chunk in chunks)
                        {
                            column.Item().Table(table =>
                            {
                                table.ColumnsDefinition(columns =>
                                {
                                    columns.RelativeColumn();
                                    columns.RelativeColumn();
                                    columns.RelativeColumn(2);
                                    columns.RelativeColumn();
                                    columns.RelativeColumn();
                                });

                                // Header Row
                                table.Header(header =>
                                {
                                    header.Cell().Background(QColors.Grey.Lighten3).Border(1).Padding(4).AlignCenter().Text("رقم السجل").SemiBold().FontSize(10);
                                    header.Cell().Background(QColors.Grey.Lighten3).Border(1).Padding(4).AlignCenter().Text("رقم التصنيف").SemiBold().FontSize(10);
                                    header.Cell().Background(QColors.Grey.Lighten3).Border(1).Padding(4).AlignCenter().Text("عنوان السجل").SemiBold().FontSize(10);
                                    header.Cell().Background(QColors.Grey.Lighten3).Border(1).Padding(4).AlignCenter().Text("نوع السجل").SemiBold().FontSize(10);
                                    header.Cell().Background(QColors.Grey.Lighten3).Border(1).Padding(4).AlignCenter().Text("وقت الإنشاء").SemiBold().FontSize(10);
                                });

                                int index = 0;

                                foreach (var item in chunk)
                                {
                                    var background = index % 2 == 0 ? QColors.White : QColors.Grey.Lighten5;

                                    table.Cell().Background(background).Border(1).BorderColor(QColors.Grey.Lighten2).Padding(4).AlignCenter().Text(item.RecordNo);
                                    table.Cell().Background(background).Border(1).BorderColor(QColors.Grey.Lighten2).Padding(4).AlignCenter().Text(item.CalssificationNo);
                                    table.Cell().Background(background).Border(1).BorderColor(QColors.Grey.Lighten2).Padding(4).Text(item.RecordTitle);
                                    table.Cell().Background(background).Border(1).BorderColor(QColors.Grey.Lighten2).Padding(4).AlignCenter().Text(item.RecordType);
                                    table.Cell().Background(background).Border(1).BorderColor(QColors.Grey.Lighten2).Padding(4).AlignCenter().Text(item.CreationTime.ToString("g"));

                                    index++;
                                }
                            });

                            // Page break بين كل 10 عناصر
                            if (chunk != chunks.Last())
                            {
                                column.Item().PageBreak();
                            }
                        }
                    });

                    // Footer
                    page.Footer().AlignCenter().Text(x =>
                    {
                        x.Span("صفحة ");
                        x.CurrentPageNumber();
                        x.Span(" من ");
                        x.TotalPages();
                    });
                });
            });

            var pdfBytes = document.GeneratePdf();
            return File(pdfBytes, "application/pdf", "SearchRecords.pdf");
        }

        public IActionResult OnGetDownloadExcel()
        {
            var data = _searchService.GetSearchModels().ToList();

            using var stream = new MemoryStream();
            using (var spreadsheetDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
                var sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Records" };
                sheets.Append(sheet);

                // Header
                var headerRow = new Row();
                headerRow.Append(
                    new Cell { CellValue = new CellValue("رقم السجل"), DataType = CellValues.String },
                    new Cell { CellValue = new CellValue("رقم التصنيف"), DataType = CellValues.String },
                    new Cell { CellValue = new CellValue("عنوان السجل"), DataType = CellValues.String },
                    new Cell { CellValue = new CellValue("نوع السجل"), DataType = CellValues.String },
                    new Cell { CellValue = new CellValue("وقت الإنشاء"), DataType = CellValues.String }
                );
                sheetData.Append(headerRow);

                // Content
                foreach (var item in data)
                {
                    var row = new Row();
                    row.Append(
                        new Cell { CellValue = new CellValue(item.RecordNo), DataType = CellValues.String },
                        new Cell { CellValue = new CellValue(item.CalssificationNo), DataType = CellValues.String },
                        new Cell { CellValue = new CellValue(item.RecordTitle), DataType = CellValues.String },
                        new Cell { CellValue = new CellValue(item.RecordType), DataType = CellValues.String },
                        new Cell { CellValue = new CellValue(item.CreationTime.ToString("g")), DataType = CellValues.String }
                    );
                    sheetData.Append(row);
                }

                workbookPart.Workbook.Save();
            }

            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "SearchRecords.xlsx");
        }

        public IActionResult OnGetDownloadWord()
        {
            var data = _searchService.GetSearchModels().ToList();

            using var stream = new MemoryStream();
            using (var wordDocument = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Wordprocessing.Document();
                var body = mainPart.Document.AppendChild(new Wordprocessing.Body());

                // Title
                var para = body.AppendChild(new Wordprocessing.Paragraph());
                var run = para.AppendChild(new Wordprocessing.Run());
                run.AppendChild(new Wordprocessing.Text("تصدير السجلات المرجعية"));

                // Table
                var table = new Wordprocessing.Table();
                
                var tableProperties = new Wordprocessing.TableProperties(
                    new Wordprocessing.TableBorders(
                        new Wordprocessing.TopBorder { Val = new EnumValue<Wordprocessing.BorderValues>(Wordprocessing.BorderValues.Single), Size = 12 },
                        new Wordprocessing.BottomBorder { Val = new EnumValue<Wordprocessing.BorderValues>(Wordprocessing.BorderValues.Single), Size = 12 },
                        new Wordprocessing.LeftBorder { Val = new EnumValue<Wordprocessing.BorderValues>(Wordprocessing.BorderValues.Single), Size = 12 },
                        new Wordprocessing.RightBorder { Val = new EnumValue<Wordprocessing.BorderValues>(Wordprocessing.BorderValues.Single), Size = 12 },
                        new Wordprocessing.InsideHorizontalBorder { Val = new EnumValue<Wordprocessing.BorderValues>(Wordprocessing.BorderValues.Single), Size = 12 },
                        new Wordprocessing.InsideVerticalBorder { Val = new EnumValue<Wordprocessing.BorderValues>(Wordprocessing.BorderValues.Single), Size = 12 }
                    )
                );
                table.AppendChild(tableProperties);

                // Header Row
                var headerRow = new Wordprocessing.TableRow();
                headerRow.Append(
                    CreateWordTableCell("رقم السجل"),
                    CreateWordTableCell("رقم التصنيف"),
                    CreateWordTableCell("عنوان السجل"),
                    CreateWordTableCell("نوع السجل"),
                    CreateWordTableCell("وقت الإنشاء")
                );
                table.AppendChild(headerRow);

                // Content
                foreach (var item in data)
                {
                    var tr = new Wordprocessing.TableRow();
                    tr.Append(
                        CreateWordTableCell(item.RecordNo),
                        CreateWordTableCell(item.CalssificationNo),
                        CreateWordTableCell(item.RecordTitle),
                        CreateWordTableCell(item.RecordType),
                        CreateWordTableCell(item.CreationTime.ToString("g"))
                    );
                    table.AppendChild(tr);
                }

                body.AppendChild(table);
            }

            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "SearchRecords.docx");
        }

        private Wordprocessing.TableCell CreateWordTableCell(string text)
        {
            var tc = new Wordprocessing.TableCell();
            tc.Append(new Wordprocessing.Paragraph(new Wordprocessing.Run(new Wordprocessing.Text(text))));
            return tc;
        }
    }
}