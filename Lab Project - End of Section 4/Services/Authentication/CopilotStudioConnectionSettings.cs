using Microsoft.Agents.CopilotStudio.Client;

namespace webchatclient.Services.Authentication
{
    internal class CopilotStudioConnectionSettings : ConnectionSettings
    {
        public string TenantId { get; }
        public string AppClientId { get; }
        public string? AppClientSecret { get; }
        public bool UseS2SConnection { get; }

        public CopilotStudioConnectionSettings(
            IConfigurationSection copilotConfig,
            IConfigurationSection azureAdConfig)
            : base(copilotConfig)
        {
            TenantId = azureAdConfig["TenantId"]
                       ?? throw new ArgumentException("TenantId not found in AzureAd config");
            AppClientId = azureAdConfig["ClientId"]
                          ?? throw new ArgumentException("ClientId not found in AzureAd config");
            AppClientSecret = azureAdConfig["ClientSecret"];
            UseS2SConnection = copilotConfig.GetValue<bool>("UseS2SConnection", false);
        }
    }
}