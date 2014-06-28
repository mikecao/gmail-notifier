using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Gmail;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Oauth2;
using Google.Apis.Oauth2.v2;
using Google.Apis.Oauth2.v2.Data;
using Google.Apis.Services;
using Google.Apis.Util;
using Google.Apis.Util.Store;

namespace GmailNotifier
{
    class GmailConnection
    {
        #region Members
        private UserCredential _credential;
        private TokenResponse _token;
        private GmailService _service;
        private string _userId;
        private int _mailCount;
        #endregion

        #region Properties
        public UserCredential Credential
        {
            get { return _credential; }
        }

        public TokenResponse Token
        {
            get { return _token; }
        }

        public GmailService Service
        {
            get { return _service; }
        }

        public int MailCount
        {
            get { return _mailCount; }
        }

        public bool IsConnected
        {
            get { return _credential !=null && _token != null && _service != null; }
        }
        #endregion

        #region Methods
        public GmailConnection(string userId)
        {
            _userId = userId;
        }

        /// <summary>
        /// Connects to the Gmail API.
        /// </summary>
        public void Connect()
        {
            using (var stream = new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
            {
                _credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    new[] {
                    GmailService.Scope.GmailReadonly
                },
                    "user",
                    CancellationToken.None,
                    GetDataStore()
                ).Result;

                _token = _credential.Token;

                _service = new GmailService(
                    new BaseClientService.Initializer()
                    {
                        ApplicationName = Properties.Resources.ApplicationName,
                        HttpClientInitializer = _credential
                    }
                );
            }
        }

        /// <summary>
        /// Gets the data storage object.
        /// </summary>
        /// <returns></returns>
        public FileDataStore GetDataStore()
        {
            return new FileDataStore(GetDataFolder());
        }

        /// <summary>
        /// Gets the data folder for storing OAuth tokens.
        /// </summary>
        /// <returns>Data folder path</returns>
        public string GetDataFolder()
        {
            return string.Format(
                @"{0}\{1}",
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                Properties.Resources.ApplicationName
            );
        }

        /// <summary>
        /// Checks if there are any unread messages in the user's inbox.
        /// </summary>
        /// <returns>Inbox status</returns>
        public bool CheckMessages()
        {
            if (_service != null)
            {
                UsersResource.MessagesResource.ListRequest request = _service.Users.Messages.List(_userId);
                request.LabelIds = new Repeatable<string>(new List<string>() { "INBOX", "UNREAD" });
                request.IncludeSpamTrash = false;

                IList<Google.Apis.Gmail.v1.Data.Message> messages = request.Execute().Messages;

                _mailCount = (messages != null) ? messages.Count : 0;

                return _mailCount > 0;
            }

            return false;
        }

        /// <summary>
        /// Disconnects from the Gmail API. Removes any stored access tokens.
        /// </summary>
        public void Disconnect()
        {
            if (_token != null)
            {
                string token = (_token.RefreshToken != null) ? _token.RefreshToken : _token.AccessToken;

                WebRequest request = WebRequest.Create("https://accounts.google.com/o/oauth2/revoke?token=" + token);

                WebResponse response = request.GetResponse();

                FileDataStore data = GetDataStore();
                data.ClearAsync();
            }
        }
        #endregion
    }
}
