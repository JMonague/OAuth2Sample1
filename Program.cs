using log4net;
using log4net.Config;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;
using Microsoft.Identity.Client;
using MimeKit;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace IMapAuthTest4
{
    class Program
    {
        private static string _clientId = ConfigurationManager.AppSettings["clientid"];
        private static string _tenantId = ConfigurationManager.AppSettings["tenantid"];
        private static string _redirUri = ConfigurationManager.AppSettings["redirecturi"];
        private static string _exchangeaccount = ConfigurationManager.AppSettings["exchangeaccount"];
        private static string _tokenCacheFile = ConfigurationManager.AppSettings["tokencachefile"];
        //private static string _authorityUri = "https://login.microsoftonline.com/{0}/oauth2/v2.0/authorize?client_id={1}&response_type=code&redirect_uri={2}&response_mode=query&scope={3}&state=12345&sso_reload=true";
        private static int _deltaDateAddMinutes = int.Parse(ConfigurationManager.AppSettings["deltaDateAddMinutes"]);

        private static readonly ILog log = LogManager.GetLogger(typeof(Program));

        public static string ClientID
        {
            get
            {
                return _clientId;
            }
        }

        public static string TenantID
        {
            get
            {
                return _tenantId;
            }
        }

        public static string RedirectUri
        {
            get
            {
                return _redirUri;
            }
        }

        /*public static string AuthorityUri
        {
            get
            {
                return string.Format(CultureInfo.InvariantCulture, _authorityUri, TenantID, ClientID,RedirectUri, "https://graph.microsoft.com/IMAP.AccessAsUser.All");
            }
        }*/
        
        public static string ExchangeAccount
        {
            get
            {
                return _exchangeaccount;
            }
        }

        /**
         * <remarks>
         * see: https://learn.microsoft.com/en-us/azure/active-directory/develop/refresh-tokens#revocation
         * </remarks>
         */
        public static string TokenCacheFile
        {
            get
            {
                return _tokenCacheFile;
            }
        }

        public static DateTime DeltaDate
        {
            get
            {
                return DateTime.UtcNow.AddMinutes(_deltaDateAddMinutes);
            }
        }

        static void Main(string[] args)
        {
            XmlConfigurator.Configure();

            using (var client = new ImapClient())
            {
                client.Connect("outlook.office365.com", 993, SecureSocketOptions.SslOnConnect);
                if (client.AuthenticationMechanisms.Contains("OAUTHBEARER") || client.AuthenticationMechanisms.Contains("XOAUTH2"))
                    AuthenticateAsync(client)
                        .GetAwaiter()
                        .GetResult();


                client.Inbox.Open(FolderAccess.ReadWrite);
                IList<UniqueId> uids = client.Inbox.Search(SearchQuery.NotSeen);
                foreach (UniqueId uid in uids)
                {
                    MimeMessage mm = client.Inbox.GetMessage(uid);
                    string msg = "Email info:";
                    msg += $"\r\nTo: {mm.To}";
                    msg += $"\r\nFrom: {mm.From}";
                    msg += $"\r\nSubject: {mm.Subject}";
                    log.Info(msg);
                }
                client.Disconnect(true);
            }
        }


        static async Task AuthenticateAsync(ImapClient client)
        {
            PublicClientApplicationOptions options = new PublicClientApplicationOptions
            {
                ClientId = Program.ClientID,
                TenantId = Program.TenantID,
                RedirectUri = Program.RedirectUri,
            };


            var publicClientApplication = PublicClientApplicationBuilder
                .CreateWithApplicationOptions(options)
                .Build();

            TokenCacheHelper.EnableSerialization(publicClientApplication.UserTokenCache);

            var scopes = new string[] {
                "email",
                "offline_access",
                "https://outlook.office.com/IMAP.AccessAsUser.All", // Only needed for IMAP
                //"https://outlook.office.com/POP.AccessAsUser.All",  // Only needed for POP
                //"https://outlook.office.com/SMTP.AccessAsUser.All", // Only needed for SMTP
            };

            if (!File.Exists(TokenCacheFile))
            {
                // For first time generation only.
                AuthenticationResult authToken = await publicClientApplication
                    .AcquireTokenInteractive(scopes)
                    .WithLoginHint(ExchangeAccount)
                    .ExecuteAsync(CancellationToken.None);

                SaslMechanism oauth2 = new SaslMechanismOAuth2(authToken.Account.Username, authToken.AccessToken);
                await client.AuthenticateAsync(oauth2);
            }
            else
            {
                log.Info($"Loading token from local cache file '{TokenCacheFile}'");

                JObject jTokenCacheFile = JObject.Parse(File.ReadAllText(TokenCacheFile));

                var jAccessToken = jTokenCacheFile["AccessToken"].First.First;
                var jRefreshToken = jTokenCacheFile["RefreshToken"].First.First;

                string accessTokenSecret = jAccessToken["secret"].ToString();
                string refreshTokenSecret = jRefreshToken["secret"].ToString();
                long accessTokenCachedAt = long.Parse(jAccessToken["cached_at"].ToString());
                long accessTokenExpiresOn = long.Parse(jAccessToken["expires_on"].ToString());
                long accessTokenExtendedExpiresOn = long.Parse(jAccessToken["extended_expires_on"].ToString());
                long accessTokenExtExpiresOn = long.Parse(jAccessToken["ext_expires_on"].ToString());

                DateTime expiryDate = DateTimeOffset.FromUnixTimeSeconds(accessTokenExpiresOn).DateTime;

                #region PrintJSonProps
                log.Info($"Cached at: {DateTimeOffset.FromUnixTimeSeconds(accessTokenCachedAt)}");
                log.Info($"expires_on: {expiryDate}");
                log.Info($"extended_expires_on: {DateTimeOffset.FromUnixTimeSeconds(accessTokenExtendedExpiresOn)}");
                log.Info($"ext_expires_on: {DateTimeOffset.FromUnixTimeSeconds(accessTokenExtExpiresOn)}");
                #endregion PrintJSonProps

                if (expiryDate.CompareTo(DeltaDate) > 0)
                {
                    SaslMechanism sasl = new SaslMechanismOAuth2("joshua.monague@salumatics.com", accessTokenSecret);
                    await client.AuthenticateAsync(sasl);
                }
                else
                {
                    //Updates the token cache file.
                    AuthenticationResult refToken = await ((IByRefreshToken)publicClientApplication).AcquireTokenByRefreshToken(scopes, refreshTokenSecret)
                        .WithTenantId(TenantID)
                        .ExecuteAsync();

                    #region PrintRefreshTokenProps
                    log.Info("Aquired token by refresh token.");
                    log.Info($"new expires_on: {refToken.ExpiresOn}");
                    log.Info($"new extended_expires_on: {refToken.ExtendedExpiresOn}");
                    log.Info($"IsExtendedLifeTimeToken? {refToken.IsExtendedLifeTimeToken}");
                    #endregion PrintRefreshTokenProps

                    SaslMechanism sasl = new SaslMechanismOAuth2(refToken.Account.Username, refToken.AccessToken);
                    await client.AuthenticateAsync(sasl);
                }
            }
        }
    }
}