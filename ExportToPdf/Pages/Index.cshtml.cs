using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using ExportToPdf.Models;
using ExportToPdf.Services;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

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
                        .FontColor(Colors.Black);

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
                                    header.Cell().Background(Colors.Grey.Lighten3).Border(1).Padding(4).AlignCenter().Text("رقم السجل").SemiBold().FontSize(10);
                                    header.Cell().Background(Colors.Grey.Lighten3).Border(1).Padding(4).AlignCenter().Text("رقم التصنيف").SemiBold().FontSize(10);
                                    header.Cell().Background(Colors.Grey.Lighten3).Border(1).Padding(4).AlignCenter().Text("عنوان السجل").SemiBold().FontSize(10);
                                    header.Cell().Background(Colors.Grey.Lighten3).Border(1).Padding(4).AlignCenter().Text("نوع السجل").SemiBold().FontSize(10);
                                    header.Cell().Background(Colors.Grey.Lighten3).Border(1).Padding(4).AlignCenter().Text("وقت الإنشاء").SemiBold().FontSize(10);
                                });

                                int index = 0;

                                foreach (var item in chunk)
                                {
                                    var background = index % 2 == 0 ? Colors.White : Colors.Grey.Lighten5;

                                    table.Cell().Background(background).Border(1).BorderColor(Colors.Grey.Lighten2).Padding(4).AlignCenter().Text(item.RecordNo);
                                    table.Cell().Background(background).Border(1).BorderColor(Colors.Grey.Lighten2).Padding(4).AlignCenter().Text(item.CalssificationNo);
                                    table.Cell().Background(background).Border(1).BorderColor(Colors.Grey.Lighten2).Padding(4).Text(item.RecordTitle);
                                    table.Cell().Background(background).Border(1).BorderColor(Colors.Grey.Lighten2).Padding(4).AlignCenter().Text(item.RecordType);
                                    table.Cell().Background(background).Border(1).BorderColor(Colors.Grey.Lighten2).Padding(4).AlignCenter().Text(item.CreationTime.ToString("g"));

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
    }
}