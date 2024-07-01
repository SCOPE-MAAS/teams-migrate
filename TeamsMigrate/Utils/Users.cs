using System;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using TeamsMigrate.ViewModels;
using System.Linq;

namespace TeamsMigrate.Utils
{
    public class Users
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(Users));

        private static Dictionary<string, string> users = new Dictionary<string, string>();

        public static List<ViewModels.SimpleUser> ScanUsers(string combinedPath)
        {
            var simpleUserList = new List<ViewModels.SimpleUser>();
            using (FileStream fs = new FileStream(combinedPath, FileMode.Open, FileAccess.Read))
            using (StreamReader sr = new StreamReader(fs))
            using (JsonTextReader reader = new JsonTextReader(sr))
            {
                while (reader.Read())
                {
                    if (reader.TokenType == JsonToken.StartObject)
                    {
                        JObject obj = JObject.Load(reader);

                        var userId = (string)obj.SelectToken("id");
                        var email = (string)obj.SelectToken("profile.email");
                        var is_bot = (bool)obj.SelectToken("is_bot");
                        var name = !is_bot ? email.Split("@")[0] : (string)obj.SelectToken("name");
                        var real_name = (string)obj.SelectToken("profile.real_name_normalized");

                        log.DebugFormat("Scanned user {0} ({1}) {2}", name, email, (string)obj.SelectToken("real_name"));

                        simpleUserList.Add(new ViewModels.SimpleUser()
                        {
                            userId = userId,
                            name = name,
                            email = email,
                            real_name = real_name,
                            is_bot = is_bot,
                        });
                    }
                }
            }

            log.InfoFormat("Total users scanned: {0}", simpleUserList.Count);
            return simpleUserList;
        }

        internal static string GetUserIdByName(string messageSender)
        {
            string principalName = messageSender + "@" + Program.CmdOptions.Domain;
            return GetUserId(principalName);
        }

        internal static string GetUserId(string id)
        {
            Helpers.httpClient.DefaultRequestHeaders.Clear();
            Helpers.httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", TeamsMigrate.Utils.Auth.AccessToken);
            Helpers.httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var url = string.Format("{0}users/{1}", O365.MsGraphEndpoint, id);
            log.Debug("GET " + url);
            var httpResponseMessage = Helpers.httpClient.GetAsync(url).Result;
            if (!httpResponseMessage.IsSuccessStatusCode)
            {
                log.DebugFormat("User '{0}' not exist", id);
                log.Debug(httpResponseMessage.Content.ReadAsStringAsync().Result.ToString());
                return "";
            }

            dynamic user = JObject.Parse(httpResponseMessage.Content.ReadAsStringAsync().Result);
            log.InfoFormat("Found user: {0}", user.id);
            return user.id;
        }

        public static string GetOrCreateId(string messageSender, List<SimpleUser> slackUserList, string domain)
        {
            try
            {
                if (users.ContainsKey(messageSender))
                {
                    return users[messageSender];
                }

                SimpleUser simpleUser = slackUserList.FirstOrDefault(w => w.name == messageSender);
                if (simpleUser == null)
                {
                    simpleUser = new SimpleUser();
                    simpleUser.real_name = messageSender;
                    simpleUser.name = messageSender;
                }

                Console.Write($"Enter the AAD object ID for Slack user '{simpleUser.real_name} ({simpleUser.name})': ");
                string aadObjectId = Console.ReadLine();

                if (string.IsNullOrEmpty(aadObjectId))
                {
                    log.Warn($"No AAD object ID provided for user '{simpleUser.real_name} ({simpleUser.name})'. This user will be skipped.");
                    return "";
                }

                users.Add(messageSender, aadObjectId);
                return aadObjectId;
            }
            catch (Exception ex)
            {
                log.Debug("Failed to get user");
                log.Debug("Failure", ex);
                return "";
            }
        }
    }
}
