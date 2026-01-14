using Microsoft.Identity.Web;
using System.Net.Http.Headers;

internal class AuthTokenHandler : DelegatingHandler
{
    private readonly IHttpContextAccessor _httpContextAccessor;
    private readonly ITokenAcquisition _tokenAcquisition;
    private readonly string _scope;
    private readonly ILogger<AuthTokenHandler> _logger;

    public AuthTokenHandler(
        IHttpContextAccessor httpContextAccessor,
        ITokenAcquisition tokenAcquisition,
        CopilotScope copilotScope,
        ILogger<AuthTokenHandler> logger)
    {
        _httpContextAccessor = httpContextAccessor;
        _tokenAcquisition = tokenAcquisition;
        _scope = copilotScope.Value;
        _logger = logger;
    }

    protected override async Task<HttpResponseMessage> SendAsync(
        HttpRequestMessage request, CancellationToken cancellationToken)
    {
        if (request.Headers.Authorization is null)
        {
            var context = _httpContextAccessor.HttpContext
                ?? throw new InvalidOperationException("No HttpContext available");

            if (context.User.Identity?.IsAuthenticated != true)
            {
                throw new InvalidOperationException("User is not authenticated");
            }

            try
            {
                var accessToken = await _tokenAcquisition
                    .GetAccessTokenForUserAsync(new[] { _scope });

                request.Headers.Authorization =
                    new AuthenticationHeaderValue("Bearer", accessToken);
            }
            catch (MicrosoftIdentityWebChallengeUserException ex)
            {
                _logger.LogWarning(ex, "Token acquisition failed - user needs to re-authenticate");
                throw new InvalidOperationException("Session expired. Please sign out and sign back in.");
            }
        }

        return await base.SendAsync(request, cancellationToken);
    }
}
