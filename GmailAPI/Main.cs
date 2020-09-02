using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using System;
using System.Data;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;

namespace GmailAPI
{
    class GmailAPI
    {
        static string[] Scopes = { GmailService.Scope.GmailReadonly };
        static string clientID = "949780797433-0jeapp019do25gqejeemfp28ud71vcso.apps.googleusercontent.com";
        static string clientSecret = "nRlcGPEmZFi66ysomXMw62jA";

        static void Main(string[] args)
        {
            // Initialise Objects
            GetEmails Emails = new GetEmails();

            List<string> emailList = Emails.getEmails(clientID, clientSecret, "leroytangyl@gmail.com", "test12345");

        }
        public static GmailService GetGmailService(string clientID, string clientSecret)
        {
            var clientSecrets = new ClientSecrets
            {
                ClientId = clientID,
                ClientSecret = clientSecret
            };

            return new GmailService(new BaseClientService.Initializer
            {
                HttpClientInitializer = GoogleWebAuthorizationBroker.AuthorizeAsync(clientSecrets, Scopes, "user", CancellationToken.None).Result
            });
        }
    }
    
}
