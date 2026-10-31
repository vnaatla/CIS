namespace Cis.Api.Auth;

/// <summary>JWT settings (configuration section <c>Auth</c>).</summary>
public class AuthOptions
{
    /// <summary>Configuration section name.</summary>
    public const string SectionName = "Auth";

    /// <summary>HMAC key (32+ characters). When empty a random key is generated, so tokens stop working after a restart.</summary>
    public string? JwtKey { get; set; }

    /// <summary>Token issuer and audience.</summary>
    public string Issuer { get; set; } = "Cis";

    /// <summary>Lifetime of a token.</summary>
    public int TokenHours { get; set; } = 10;
}
