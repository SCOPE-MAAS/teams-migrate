using System;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Threading;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using TeamsMigrate.Models;
using System.Linq;

namespace TeamsMigrate.Utils
{
    public class Channels
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(Channels));

        public static List<Slack.Channels> ScanSlackChannelsJson(string combinedPath, string membershipType = "standard")
        {
            List<Slack.Channels> slackChannels = new List<Slack.Channels>();

            using (FileStream fs = new FileStream(combinedPath, FileMode.Open, FileAccess.Read))
            using (StreamReader sr = new StreamReader(fs))
            using (JsonTextReader reader = new JsonTextReader(sr))
            {
                while (reader.Read())
                {
                    if (reader.TokenType == JsonToken.StartObject)
                    {
                        JObject obj = JObject.Load(reader);

                        var channelId = (string)obj.SelectToken("id");
                        if (channelId == null)
                        {
                            channelId = "";
                        }

                        slackChannels.Add(new Models.Slack.Channels()
                        {
                            channelId = channelId,
                            channelName = obj["name"].ToString(),
                            channelDescription = obj["purpose"]["value"].ToString(),
                            membershipType = membershipType,
                            members = obj["members"]?.ToObject<List<string>>() ?? new List<string>()
                        });
                    }
                }
            }
            return slackChannels;
        }

        internal static void DeleteChannel(string selectedTeamId, string channelId)
        {
            Helpers.httpClient.DefaultRequestHeaders.Clear();
            Helpers.httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", TeamsMigrate.Utils.Auth.AccessToken);
            Helpers.httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var url = String.Format("{0}teams/{1}/channels/{2}", O365.MsGraphEndpoint, selectedTeamId, channelId);
            var httpResponseMessage = Helpers.httpClient.DeleteAsync(url).Result;
            if (httpResponseMessage.IsSuccessStatusCode)
            {
                log.DebugFormat("Channel {0} in team {1} has been deleted", channelId, selectedTeamId);
            }
            else
            {
                log.DebugFormat("Failed to delete channel {0} in team {1}", channelId, selectedTeamId);
                log.Debug(httpResponseMessage.Content.ReadAsStringAsync().Result);
            }
        }

        internal static void AddChannel(string selectedTeamId, string channelName, string description)
        {
            var createChannelUrl = String.Format("{0}teams/{1}/channels", O365.MsGraphEndpoint, selectedTeamId);
            dynamic newChannelObject = new JObject();
            newChannelObject.displayName = channelName;
            newChannelObject.description = description;
            newChannelObject.Add("@odata.type", "#microsoft.graph.channel");

            var createChannelPostData = JsonConvert.SerializeObject(newChannelObject);
            log.DebugFormat("POST {0} \n{1}", createChannelUrl, createChannelPostData);

            Helpers.httpClient.DefaultRequestHeaders.Clear();
            Helpers.httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", TeamsMigrate.Utils.Auth.AccessToken);
            Helpers.httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var httpResponseMessage = Helpers.httpClient.PostAsync(createChannelUrl, new StringContent(createChannelPostData, Encoding.UTF8, "application/json")).Result;

            if (!httpResponseMessage.IsSuccessStatusCode)
            {
                log.Error("Channel could not be created");
                log.Debug(httpResponseMessage.Content.ReadAsStringAsync().Result);
            }
        }

        internal static void CompleteTeamMigration(string selectedTeamId)
        {
            var completeTeamMigrationUrl = String.Format("{0}teams/{1}/completeMigration", O365.MsGraphBetaEndpoint, selectedTeamId);

            dynamic completeTeamMigrationObject = new JObject();
            completeTeamMigrationObject.Add("@odata.type", "#microsoft.graph.completeMigration");

            var completeTeamMigrationPostData = JsonConvert.SerializeObject(completeTeamMigrationObject);
            log.DebugFormat("POST {0} \n{1}", completeTeamMigrationUrl, completeTeamMigrationPostData);

            Helpers.httpClient.DefaultRequestHeaders.Clear();
            Helpers.httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", TeamsMigrate.Utils.Auth.AccessToken);
            Helpers.httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var httpResponseMessage = Helpers.httpClient.PostAsync(completeTeamMigrationUrl, new StringContent(completeTeamMigrationPostData, Encoding.UTF8, "application/json")).Result;

            if (!httpResponseMessage.IsSuccessStatusCode)
            {
                log.Error("Complete team migration could not be finished");
                log.Debug(httpResponseMessage.Content.ReadAsStringAsync().Result);
            }
        }

        internal static void AssignTeamOwnerships(string selectedTeamId)
        {
            var authHelper = new Utils.O365.AuthenticationHelper() { AccessToken = TeamsMigrate.Utils.Auth.AccessToken };
            Microsoft.Graph.GraphServiceClient gcs = new Microsoft.Graph.GraphServiceClient(authHelper);

            var owners = gcs.Groups[selectedTeamId].Owners.Request().GetAsync().Result;

            foreach (var owner in owners)
            {
                log.DebugFormat("Adding owner: {0}", owner.Id);
                TeamsMigrate.Utils.Users.AddOwner(selectedTeamId, owner.Id);
            }
        }

        internal static void AssignChannelsMembership(string selectedTeamId, List<Models.Combined.ChannelsMapping> channelsMapping, List<ViewModels.SimpleUser> slackUserList)
        {
            foreach (var channel in channelsMapping)
            {
                var authHelper = new Utils.O365.AuthenticationHelper() { AccessToken = TeamsMigrate.Utils.Auth.AccessToken };
                Microsoft.Graph.GraphServiceClient gcs = new Microsoft.Graph.GraphServiceClient(authHelper);

                foreach (var member in channel.members)
                {
                    var user = slackUserList.FirstOrDefault(u => u.userId == member);
                    if (user != null)
                    {
                        var userId = Users.GetUserIdByName(user.name);
                        if (!string.IsNullOrEmpty(userId))
                        {
                            Users.AddMemberChannel(selectedTeamId, channel.id, userId);
                        }
                    }
                }
            }
        }
    }
}
