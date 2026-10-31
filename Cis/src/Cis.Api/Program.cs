using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application;
using Cis.Application.Auth.Login;
using Cis.Infrastructure;
using Cis.Infrastructure.Data;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.IdentityModel.Tokens;
using Microsoft.OpenApi;
using Serilog;

var builder = WebApplication.CreateBuilder(args);

// Centralized logging: console + rolling files, configured in the "Serilog" section of appsettings.json.
builder.Host.UseSerilog((context, services, logging) => logging.ReadFrom.Configuration(context.Configuration).Enrich.FromLogContext());

// Composition root: application (use cases) + infrastructure (adapters) + the API's own pieces (JWT issuer, filters, error handling).
builder.Services.AddCisApplication();
builder.Services.AddCisInfrastructure(builder.Configuration);
builder.Services.Configure<AuthOptions>(builder.Configuration.GetSection(AuthOptions.SectionName));
builder.Services.AddSingleton<TokenService>();
builder.Services.AddSingleton<ITokenIssuer>(sp => sp.GetRequiredService<TokenService>());
builder.Services.AddProblemDetails();
builder.Services.AddExceptionHandler<ApiExceptionHandler>();
builder.Services.AddHealthChecks().AddCheck<DatabaseHealthCheck>("database");

builder.Services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme).AddJwtBearer();
builder.Services.AddOptions<JwtBearerOptions>(JwtBearerDefaults.AuthenticationScheme)
    .Configure<TokenService>((options, tokens) =>
    {
        options.MapInboundClaims = false;
        options.TokenValidationParameters = new TokenValidationParameters
        {
            ValidIssuer = tokens.Issuer,
            ValidAudience = tokens.Issuer,
            IssuerSigningKey = tokens.SigningKey,
            NameClaimType = System.Security.Claims.ClaimTypes.Name,
        };
    });
builder.Services.AddAuthorization(options =>
    options.FallbackPolicy = new Microsoft.AspNetCore.Authorization.AuthorizationPolicyBuilder().RequireAuthenticatedUser().Build());

var corsOrigins = builder.Configuration.GetSection("Cors:Origins").Get<string[]>() ?? [];
builder.Services.AddCors(o => o.AddDefaultPolicy(p => p.WithOrigins(corsOrigins).AllowAnyHeader().AllowAnyMethod()));

builder.Services.AddControllers(options => options.Filters.Add<ValidationFilter>())
    // Enums travel as their names ("Text", "Number", "Pdf" ...), which is what the Angular app compares against.
    .AddJsonOptions(options => options.JsonSerializerOptions.Converters.Add(new System.Text.Json.Serialization.JsonStringEnumConverter()));
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(options =>
{
    var xmlDocs = Path.Combine(AppContext.BaseDirectory, "Cis.Api.xml");
    if (File.Exists(xmlDocs)) options.IncludeXmlComments(xmlDocs);
    options.CustomSchemaIds(t => (t.FullName ?? t.Name).Replace("Cis.Api.", "").Replace('+', '.'));
    options.AddSecurityDefinition("Bearer", new OpenApiSecurityScheme
    {
        Name = "Authorization",
        Type = SecuritySchemeType.Http,
        Scheme = "bearer",
        BearerFormat = "JWT",
        In = ParameterLocation.Header,
        Description = "Paste the token from POST /api/auth/login.",
    });
    options.AddSecurityRequirement(document => new OpenApiSecurityRequirement { [new OpenApiSecuritySchemeReference("Bearer", document)] = [] });
});


var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseSerilogRequestLogging(options =>
    options.EnrichDiagnosticContext = (diagnostics, http) => diagnostics.Set("User", http.User.Identity?.Name ?? "anonymous"));
app.UseExceptionHandler();
app.UseStatusCodePages();

app.UseDefaultFiles();
app.UseStaticFiles();
app.UseCors();
app.UseAuthentication();
app.UseAuthorization();
app.MapControllers();
app.MapHealthChecks("/api/health").AllowAnonymous();
app.MapFallbackToFile("index.html").AllowAnonymous();

app.Run();

/// <summary>Entry point (public for WebApplicationFactory in the tests).</summary>
public partial class Program;
