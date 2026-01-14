using Microsoft.Agents.CopilotStudio.Client;
using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.Server;
using Microsoft.AspNetCore.DataProtection;
using Microsoft.Extensions.AI;
using Microsoft.Identity.Web;
using Microsoft.Identity.Web.UI;
using webchatclient.Components;
using webchatclient.Services;
using webchatclient.Services.Authentication;

var builder = WebApplication.CreateBuilder(args);

// Add Razor components with interactive server-side rendering
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

builder.Services.AddDataProtection()
    .UseEphemeralDataProtectionProvider();

// Build connection settings
var copilotSettings = new CopilotStudioConnectionSettings(
    builder.Configuration.GetSection("CopilotStudio"),
    builder.Configuration.GetSection("AzureAd"));

string copilotScope = CopilotClient.ScopeFromSettings(copilotSettings);

builder.Services.AddHttpContextAccessor();

// Configure authentication with MSAL using in memory cache
builder.Services.AddAuthentication(OpenIdConnectDefaults.AuthenticationScheme)
    .AddMicrosoftIdentityWebApp(builder.Configuration.GetSection("AzureAd"))
    .EnableTokenAcquisitionToCallDownstreamApi(new[] { copilotScope })
    .AddInMemoryTokenCaches();


// Add offline_access to get refresh tokens
builder.Services.Configure<OpenIdConnectOptions>(OpenIdConnectDefaults.AuthenticationScheme, options =>
{
    options.Scope.Add("offline_access");
});

// Add controllers with Microsoft Identity UI
builder.Services.AddControllersWithViews()
    .AddMicrosoftIdentityUI();

// Add authorization
builder.Services.AddAuthorization();
builder.Services.AddCascadingAuthenticationState();
builder.Services.AddScoped<AuthenticationStateProvider, ServerAuthenticationStateProvider>();

builder.Services.AddSingleton(copilotSettings);
builder.Services.AddSingleton(new CopilotScope(copilotScope));


// Register HttpClient for Copilot Studio with token handler
builder.Services.AddScoped<AuthTokenHandler>();
builder.Services.AddHttpClient("mcs")
    .AddHttpMessageHandler<AuthTokenHandler>();

// Register CopilotClient
builder.Services.AddScoped<CopilotClient>(sp =>
{
    var logger = sp.GetRequiredService<ILoggerFactory>().CreateLogger<CopilotClient>();
    return new CopilotClient(copilotSettings, sp.GetRequiredService<IHttpClientFactory>(), logger, "mcs");
});

// Register CopilotStudioIChatClient
builder.Services.AddScoped<CopilotStudioIChatClient>(sp =>
{
    var copilotClient = sp.GetRequiredService<CopilotClient>();
    return new CopilotStudioIChatClient(copilotClient);
});


builder.Services.AddScoped<IChatClient>(sp => sp.GetRequiredService<CopilotStudioIChatClient>());

var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseAuthentication();
app.UseAuthorization();
app.UseAntiforgery();
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();
app.Run();

public record CopilotScope(string Value);