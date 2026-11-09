using Cis.Application.Payment;

namespace Cis.Api.Tests;

/// <summary>ConvertCurrencyToEnglish as the desktop.</summary>
public class AmountInWordsTests
{
    [Theory]
    [InlineData("15", "Fifteen Only")]
    [InlineData("1", "One Dollar Only")]
    [InlineData("105", "One Hundred Five Only")]
    [InlineData("115", "One Hundred  And Fifteen Only")]
    [InlineData("1250.50", "One Thousand Two Hundred  And Fifty And Cents Fifty  Only")]
    [InlineData("65.05", "Sixty Five And Cents Five Only")]
    public void Wording_matches_the_desktop(string amount, string expected) =>
        Assert.Equal(expected, AmountInWords.Wording(decimal.Parse(amount, System.Globalization.CultureInfo.InvariantCulture)));

    [Fact]
    public void Long_wording_is_split_into_dollars_and_cents()
    {
        var (one, two) = AmountInWords.Lines(123456.78m);
        Assert.Equal("One Hundred  And Twenty Three Thousand Four Hundred  And Fifty Six", one);
        Assert.Equal(" And Cents Seventy Eight Only", two);
        // As stored for PV CP100468765 (7,254.15) by the desktop.
        Assert.Equal(("Seven Thousand Two Hundred  And Fifty Four", " And Cents Fifteen Only"), AmountInWords.Lines(7254.15m));
    }
}
