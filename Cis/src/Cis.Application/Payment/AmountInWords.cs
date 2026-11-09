using System.Globalization;

namespace Cis.Application.Payment;

/// <summary>CurrencyConversion.bas ConvertCurrencyToEnglish and the cheque wording of frmPaymentCLNInvoice / frmPaymentMemberAuto08, ported as is
/// (including its spacing, e.g. "One Hundred  And Fifteen").</summary>
public static class AmountInWords
{
    private static readonly string[] Place = ["", "", " Thousand ", " Million ", " Billion ", " Trillion ", "", "", "", ""];

    /// <summary>ConvertCurrencyToEnglish: the dollars and cents parts and their concatenation.</summary>
    public static (string Words, string Dollars, string Cents) Convert(decimal amount)
    {
        // VB Str(): the shortest decimal text without trailing zeros.
        var number = amount.ToString("0.############", CultureInfo.InvariantCulture);
        var cents = "";
        var point = number.IndexOf('.');
        if (point >= 0)
        {
            cents = Tens((number[(point + 1)..] + "00")[..2]);
            number = number[..point].Trim();
        }
        if (number.StartsWith('-')) number = number[1..];
        var dollars = "";
        var count = 1;
        while (number.Length > 0)
        {
            var temp = Hundreds(number.Length > 3 ? number[^3..] : number);
            if (temp.Length > 0) dollars = temp + Place[Math.Min(count, 9)] + dollars;
            number = number.Length > 3 ? number[..^3] : "";
            count++;
        }
        if (dollars == "One") dollars = "One Dollar";
        cents = cents switch { "" => "", "One" => " And  Cent One", _ => " And Cents " + cents };
        return (dollars + cents, dollars, cents);
    }

    /// <summary>txtChequeAmt_Change: the amount wording shown on the screen ("... Only", a leading "And " removed).</summary>
    public static string Wording(decimal amount)
    {
        var words = Convert(amount).Words;
        if (words.StartsWith("And", StringComparison.Ordinal)) words = words.Length > 4 ? words[4..] : "";
        return words + " Only";
    }

    /// <summary>UpdateRecord: PVWordingOne / PVWordingTwo (over 55 characters the dollars and the cents go on separate lines).</summary>
    public static (string One, string Two) Lines(decimal amount)
    {
        var wording = Wording(amount);
        if (wording.Length <= 55) return (wording.Trim(), "");
        var (_, dollars, cents) = Convert(amount);
        var one = dollars.StartsWith("And", StringComparison.Ordinal) ? (dollars.Length > 4 ? dollars[4..] : "") : dollars;
        return (one, cents.Trim().Length > 0 ? cents + " Only" : "");
    }

    private static string Hundreds(string n)
    {
        if (!int.TryParse(n, out var v) || v == 0) return "";
        n = ("000" + n)[^3..];
        var result = n[0] != '0' ? Digit(n[0]) + " Hundred " : "";
        result += n[1] != '0' ? " And " + Tens(n[1..]) : Digit(n[2]);
        return result.Trim();
    }

    private static string Tens(string t)
    {
        if (t.Length > 0 && t[0] == '1')
            return int.Parse(t, CultureInfo.InvariantCulture) switch
            {
                10 => "Ten", 11 => "Eleven", 12 => "Twelve", 13 => "Thirteen", 14 => "Fourteen", 15 => "Fifteen", 16 => "Sixteen", 17 => "Seventeen", 18 => "Eighteen",
                19 => "Nineteen", _ => "",
            };
        var result = (t.Length > 0 ? t[0] : '0') switch
        {
            '2' => "Twenty ", '3' => "Thirty ", '4' => "Forty ", '5' => "Fifty ", '6' => "Sixty ", '7' => "Seventy ", '8' => "Eighty ", '9' => "Ninety ", _ => "",
        };
        return result + Digit(t.Length > 0 ? t[^1] : '0');
    }

    private static string Digit(char d) => d switch
    {
        '1' => "One", '2' => "Two", '3' => "Three", '4' => "Four", '5' => "Five", '6' => "Six", '7' => "Seven", '8' => "Eight", '9' => "Nine", _ => "",
    };
}
