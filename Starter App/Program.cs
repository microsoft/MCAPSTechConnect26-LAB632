using Microsoft.Extensions.AI;
using webchatclient.Components;
using webchatclient.Services;

var builder = WebApplication.CreateBuilder(args);

// Add Razor components with interactive server-side rendering
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

// Register CopilotStudioIChatClient
builder.Services.AddScoped<CopilotStudioIChatClient>(sp =>
{
    return new CopilotStudioIChatClient();
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
app.UseAntiforgery();
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();
app.Run();

public record CopilotScope(string Value);