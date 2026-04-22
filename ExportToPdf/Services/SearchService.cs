using ExportToPdf.Models;

namespace ExportToPdf.Services;

public class SearchService
{
    public IEnumerable<SearchModel> GetSearchModels()
    {
        var list = new List<SearchModel>();
        for (int i = 1; i <= 100; i++)
        {
            list.Add(new SearchModel
            {
                RecordNo = $"REC-{(i).ToString("D4")}",
                CalssificationNo = $"CLASS-{i}",
                RecordTitle = $"Search Record Title {i}",
                RecordType = i % 2 == 0 ? "Type A" : "Type B",
                CreationTime = DateTime.Now.AddDays(-i)
            });
        }
        return list;
    }
}
