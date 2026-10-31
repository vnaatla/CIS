using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Extensions.Options;
using Microsoft.IdentityModel.Tokens;

namespace Cis.Api.Auth;

/// <summary>Issues the JWT that carries the user's code, name and USR.USRAccess flags.</summary>
public class TokenService
{
    /// <summary>Claim holding USR.USRAccess.</summary>
    public const string AccessClaim = "usr_access";

    private readonly AuthOptions _options;

    /// <summary>Creates the service; a random signing key is used when <see cref="AuthOptions.JwtKey"/> is empty.</summary>
    public TokenService(IOptions<AuthOptions> options)
    {
        _options = options.Value;
        SigningKey = new SymmetricSecurityKey(string.IsNullOrWhiteSpace(_options.JwtKey)
            ? RandomNumberGenerator.GetBytes(64)
            : Encoding.UTF8.GetBytes(_options.JwtKey));
    }

    /// <summary>Key the tokens are signed and validated with.</summary>
    public SymmetricSecurityKey SigningKey { get; }

    /// <summary>Issuer and audience of the tokens.</summary>
    public string Issuer => _options.Issuer;

    /// <summary>A signed token for the user and its expiry (UTC).</summary>
    public (string Token, DateTime ExpiresUtc) Create(string userCode, string userName, string access)
    {
        var expires = DateTime.UtcNow.AddHours(_options.TokenHours);
        var claims = new[]
        {
            new Claim(JwtRegisteredClaimNames.Sub, userCode),
            new Claim(ClaimTypes.Name, userCode),
            new Claim("name", userName),
            new Claim(AccessClaim, access),
        };
        var token = new JwtSecurityToken(_options.Issuer, _options.Issuer, claims, expires: expires,
            signingCredentials: new SigningCredentials(SigningKey, SecurityAlgorithms.HmacSha256));
        return (new JwtSecurityTokenHandler().WriteToken(token), expires);
    }
}
