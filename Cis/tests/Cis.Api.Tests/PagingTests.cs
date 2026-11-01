using Cis.Application.Common;

namespace Cis.Api.Tests;

public class PagingTests
{
    private sealed record Item(string Code, decimal Amount);

    [Fact]
    public void Cuts_the_requested_page_and_keeps_the_total()
    {
        var rows = Enumerable.Range(1, 120).Select(i => new Item($"C{i:000}", i)).ToList();
        var page = Paging.Apply(rows, new PageQuery(2, 50));
        Assert.Equal(120, page.Total);
        Assert.Equal(20, page.Rows.Count);
        Assert.Equal("C101", page.Rows[0].Code);
        Assert.Equal(1000, Paging.Apply(Enumerable.Range(1, 1500).Select(i => new Item("x", i)).ToList(), new PageQuery(0, 5000)).Rows.Count);
    }

    [Fact]
    public void Sorts_records_and_dictionary_rows_by_key_before_paging()
    {
        var items = new List<Item> { new("b", 10), new("a", 2), new("c", 100) };
        Assert.Equal(["c", "b", "a"], Paging.Apply(items, new PageQuery(0, 50, "amount", true)).Rows.Select(r => r.Code));
        var dicts = items.Select(i => new Dictionary<string, object?> { ["PLNCode"] = i.Code, ["AnnLimit"] = i.Amount }).ToList();
        Assert.Equal(["a", "b", "c"], Paging.Apply(dicts, new PageQuery(0, 2, "plncode")).Rows.Select(r => (string)r["PLNCode"]!).Append("c").Distinct());
        Assert.Equal(2m, Paging.Apply(dicts, new PageQuery(0, 1, "AnnLimit")).Rows[0]["AnnLimit"]);
    }
}
