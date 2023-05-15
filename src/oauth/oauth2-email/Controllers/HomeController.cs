using Microsoft.AspNetCore.Mvc;
using oauth2_email.Models;
using System.Diagnostics;
using System.Net.Http.Headers;
using Azure.Identity;
using oauth2_email.Helpers;
using Microsoft.Graph;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;

namespace oauth2_email.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        private readonly IConfiguration _configuration;

        private static Lazy<GraphServiceClient> _graphServiceClient;

        private static string? _accessToken;

        private static User? _user;

        public HomeController(ILogger<HomeController> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
            _graphServiceClient = new Lazy<GraphServiceClient>(() => CreateGraphServiceClient(_accessToken));
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        [Route("oauth/start")]
        public IActionResult StartOauthFlow()
        {
            var uriBuilder = new UriBuilder("https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize");
            uriBuilder.AddToQuery("scope", FormatScopes());
            uriBuilder.AddToQuery("state", Guid.NewGuid().ToString());
            uriBuilder.AddToQuery("response_type", "code");
            uriBuilder.AddToQuery("response_mode", value: "query");
            uriBuilder.AddToQuery("prompt", "login");
            uriBuilder.AddToQuery("client_id", _configuration["ClientId"]);
            uriBuilder.AddToQuery("client_secret", _configuration["ClientSecret"]);
            uriBuilder.AddToQuery("redirect_uri", _configuration["RedirectUri"]);

            var uri = uriBuilder.ToString();
            return Redirect(uri);
        }

        [HttpPost]
        public async Task<IActionResult> SendEmail(SendEmailModel model)
        {
            var message = new Message
            {
                Subject = model.Subject,
                Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = model.Body
                },
                ToRecipients = new List<Recipient>
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = model.To
                        }
                    }
                }
            };

            await _graphServiceClient.Value.Me
                .SendMail(message, SaveToSentItems: true)
                .Request()
                .PostAsync();


            model.Subject = string.Empty;
            model.To = string.Empty;
            model.Body = string.Empty;
            model.IsSent = true;

            return View(model);
        }

        [HttpGet]
        [Route("oauth/complete")]
        public async Task<IActionResult> CompeteOauthFlow(string code)
        {
            var rest = new RestClient("https://login.microsoftonline.com/organizations");
            var exchangeCodeRequest = new RestRequest("oauth2/v2.0/token") { Method = Method.Post };
            exchangeCodeRequest.AddParameter("grant_type", "authorization_code");
            exchangeCodeRequest.AddParameter("code", code);
            exchangeCodeRequest.AddParameter("redirect_uri", _configuration["RedirectUri"]);
            exchangeCodeRequest.AddParameter("client_id", _configuration["ClientId"]);
            exchangeCodeRequest.AddParameter("client_secret", _configuration["ClientSecret"]);

            var exchangeCodeResponse = await rest.ExecutePostAsync(exchangeCodeRequest);
            var json = JObject.Parse(exchangeCodeResponse.Content);
            var accessToken = json["access_token"].Value<string>();
            var refreshToken = json["refresh_token"].Value<string>();
            var expiresInSeconds = json["expires_in"].Value<int>();
            var tokenExpiration = DateTime.UtcNow.AddSeconds(expiresInSeconds);

            var exchangeTokenRequest = new RestRequest("oauth2/v2.0/token") { Method = Method.Post };
            exchangeTokenRequest.AddParameter("grant_type", "refresh_token");
            exchangeTokenRequest.AddParameter("refresh_token", refreshToken);
            exchangeTokenRequest.AddParameter("redirect_uri", _configuration["RedirectUri"]);
            exchangeTokenRequest.AddParameter("client_id", _configuration["ClientId"]);
            exchangeTokenRequest.AddParameter("client_secret", _configuration["ClientSecret"]);

            var exchangeTokenResponse = await rest.ExecutePostAsync(exchangeTokenRequest);
            json = JObject.Parse(exchangeTokenResponse.Content);
            var newAccessToken = json["access_token"].Value<string>();
            var newRefreshToken = json["refresh_token"].Value<string>(); //should be the same as the old one.
            expiresInSeconds = json["expires_in"].Value<int>();
            tokenExpiration = DateTime.UtcNow.AddSeconds(expiresInSeconds);

            _accessToken = newAccessToken;

            _user = await _graphServiceClient.Value.Me.Request().GetAsync();
            
            var filter = $"createdDateTime ge {ToJsonDateTime(DateTime.UtcNow.Date.AddDays(-365))} and createdDateTime le {ToJsonDateTime(DateTime.UtcNow.Date.AddDays(1))}";

            var pagedEmails = await _graphServiceClient.Value.Me.MailFolders.Inbox.Messages.Request()
                .Filter(filter)
                .OrderBy("createdDateTime")
                .GetAsync();

            var model = new SendEmailModel
            {
                Body = string.Empty,
                IsSent = false,
                Subject = string.Empty,
                SignedInAs = _user.DisplayName,
                IsSignedIn = true,
                To = string.Empty,
                Emails = new List<Message>()
            };

            var pageIterator = PageIterator<Message>.CreatePageIterator(_graphServiceClient.Value, pagedEmails,
                message =>
                {
                    model.Emails.Add(message);
                    return true;
                });

            await pageIterator.IterateAsync();

            return View("SendEmail", model);
        }

        private static string FormatScopes()
        {
            var scopes = new List<string>
            {
                "User.Read",
                "Mail.Send",
                "Mail.Read",
                "Mail.Read.Shared",
                "Mail.ReadBasic",
                "offline_access"
            };

            return string.Join(' ', scopes).Trim();
        }

        public static string ToJsonDateTime(DateTime dateTime)
        {
            return JsonConvert.ToString(dateTime).Trim('\"');
        }

        private static GraphServiceClient CreateGraphServiceClient(string accessToken)
        {
            var delegateAuthProvider = new DelegateAuthenticationProvider(requestMessage =>
            {
                requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);
                return Task.FromResult(0);
            });

            var graphClient = new GraphServiceClient(delegateAuthProvider);

            return graphClient;
        }
    }
}
