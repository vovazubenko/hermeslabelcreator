using Newtonsoft.Json;

namespace HermesLabelCreator.Services
{
    public class AuthorizationResponse
    {
        [JsonProperty("access_token")]
        public string AccessToken { get; set; }
    }
}
