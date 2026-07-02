using Xunit;

namespace Shinobi.Fuda.Tests;

public class TemplateDataTests
{
    private sealed class SimpleRow
    {
        public string Name { get; set; } = "";
        public int Qty { get; set; }
    }

    private sealed class CardLike
    {
        public string Number { get; set; } = "";
        public List<SimpleRow> Lines { get; set; } = new();
    }

    private sealed class RootDto
    {
        public string CustomerName { get; set; } = "";
        public List<SimpleRow> Records { get; set; } = new();
        public List<CardLike> Cards { get; set; } = new();

        [TemplateIgnore]
        public string Secret { get; set; } = "should not appear";
    }

    [Fact]
    public void FlattensScalarsFromRootProperties()
    {
        var data = TemplateData.FromDto(new RootDto { CustomerName = "Jane" });

        Assert.Equal("Jane", data.Scalars["CustomerName"]);
    }

    [Fact]
    public void RespectsTemplateIgnoreAttribute()
    {
        var data = TemplateData.FromDto(new RootDto { CustomerName = "Jane" });

        Assert.False(data.Scalars.ContainsKey("Secret"));
    }

    [Fact]
    public void FlatListOfScalarOnlyItemsBecomesCollection()
    {
        var dto = new RootDto
        {
            Records = new List<SimpleRow> { new() { Name = "Apple", Qty = 5 } },
        };

        var data = TemplateData.FromDto(dto);

        Assert.True(data.Collections.ContainsKey("Records"));
        Assert.False(data.Blocks.ContainsKey("Records"));
        Assert.Equal("Apple", data.Collections["Records"][0]["Name"]);
        Assert.Equal("5", data.Collections["Records"][0]["Qty"]);
    }

    [Fact]
    public void ListOfItemsWithTheirOwnListsBecomesBlock()
    {
        var dto = new RootDto
        {
            Cards = new List<CardLike>
            {
                new() { Number = "1", Lines = new List<SimpleRow> { new() { Name = "x", Qty = 1 } } },
            },
        };

        var data = TemplateData.FromDto(dto);

        Assert.True(data.Blocks.ContainsKey("Cards"));
        Assert.False(data.Collections.ContainsKey("Cards"));

        var cardItem = data.Blocks["Cards"][0];
        Assert.Equal("1", cardItem.Scalars["Number"]);
        Assert.True(cardItem.Collections.ContainsKey("Lines"));
        Assert.Equal("x", cardItem.Collections["Lines"][0]["Name"]);
    }

    [Fact]
    public void NullListBecomesEmptyCollectionNotAnError()
    {
        var dto = new RootDto { Records = null! };

        var data = TemplateData.FromDto(dto);

        Assert.False(data.Collections.ContainsKey("Records"));
        Assert.False(data.Scalars.ContainsKey("Records"));
    }
}
