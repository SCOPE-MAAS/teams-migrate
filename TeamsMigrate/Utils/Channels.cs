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
using static TeamsMigrate.Models.MsTeams;
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
                            members = obj["members"].ToObject<List<string>>()
                        });
                    }
                }
            }
            return slackChannels;
        }

        public static string GetTeamIdByName(string teamName)
        {
            Helpers.httpClient.DefaultRequestHeaders.Clear();
            Helpers.httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", TeamsMigrate.Utils.Auth.AccessToken);
            Helpers.httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var url = String.Format("{0}groups/?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')", O365.MsGraphBetaEndpoint);
            var httpResponseMessage = Helpers.httpClient.GetAsync(url).Result;

            if (httpResponseMessage.IsSuccessStatusCode)
            {
                var teams = JsonConvert.DeserializeObject<Models.MsTeams.Team>(httpResponseMessage.Content.ReadAsStringAsync().Result);
                var team = teams.value.FirstOrDefault(t => t.displayName.Equals(teamName, StringComparison.CurrentCultureIgnoreCase));
                if (team != null)
                {
                    return team.id;
                }
            }

            return null;
        }

        public static List<Combined.ChannelsMapping> GetChannelMappings(List<Slack.Channels> slackChannels)
        {
            var combinedChannelsMapping = new List<Combined.ChannelsMapping>();

            foreach (var slackChannel in slackChannels)
            {
                Console.Write($"Enter the name of the corresponding Teams channel for Slack channel '{slackChannel.channelName}': ");
                string teamsChannelName = Console.ReadLine();

                var existingTeamsChannel = GetExistingChannelByName(teamsChannelName);
                if (existingTeamsChannel != null)
                {
                    combinedChannelsMapping.Add(new Combined.ChannelsMapping()
                    {
                        id = existingTeamsChannel.id,
                        displayName = existingTeamsChannel.displayName,
                        description = existingTeamsChannel.description,
                        slackChannelId = slackChannel.channelId,
                        slackChannelName = slackChannel.channelName,
                        folderId = "",
                        members = new List<string>(slackChannel.members)
                    });
                }
                else
                {
                    log.WarnFormat("Teams channel '{0}' does not exist. Skipping mapping for Slack channel '{1}'.", teamsChannelName, slackChannel.channelName);
                }
            }

            return combinedChannelsMapping;
        }

        public static MsTeams.Channel GetExistingChannelByName(string channelName)
        {
            Helpers.httpClient.DefaultRequestHeaders.Clear();
            Helpers.httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", TeamsMigrate.Utils.Auth.AccessToken);
            Helpers.httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var url = String.Format("{0}teams/{1}/channels", O365.MsGraphBetaEndpoint, channelName);
            var httpResponseMessage = Helpers.httpClient.GetAsync(url).Result;

            if (httpResponseMessage.IsSuccessStatusCode)
            {
                var msTeamsTeam = JsonConvert.DeserializeObject<MsTeams.Team>(httpResponseMessage.Content.ReadAsStringAsync().Result);
                return msTeamsTeam.value.FirstOrDefault(t => t.displayName.Equals(channelName, StringComparison.CurrentCultureIgnoreCase));
            }

            return null;
        }

        internal static void CompleteTeamMigration(string selectedTeamId)
        {
            if (Program.CmdOptions.ReadOnly)
            {
                log.Debug("skip operation due to readonly mode");
                return;
            }
            var channels = GetExistingChannelsInMsTeams(selectedTeamId);
            int i = 1;
            using (var progress = new ProgressBar("Complete migration"))
            {
                foreach (Channel channel in channels)
                {
                    CompleteChannelMigration(selectedTeamId, channel.id);
                    progress.Report((double)i++ / channels.Count);
                }
            }

            Helpers.httpClient.DefaultRequestHeaders.Clear();
            Helpers.httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", TeamsMigrate.Utils.Auth.AccessToken);
            Helpers.httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            log.Debug("POST " + O365.MsGraphBetaEndpoint + "teams/" + selectedTeamId + "/completeMigration");

            var httpResponseMessage = Helpers.httpClient.PostAsync(O365.MsGraphBetaEndpoint + "teams/" + selectedTeamId + "/completeMigration", new StringContent("", Encoding.UTF8, "application/json")).Result;

            if (!httpResponseMessage.IsSuccessStatusCode)
            {
                log.Error("Failed to complete team migration");
                log.Debug(httpResponseMessage.Content.ReadAsStringAsync().Result);
            }
        }

        internal static void CompleteChannelMigration(string selectedTeamId, string channelId)
        {
            try
            {
                Helpers.httpClient.DefaultRequestHeaders.Clear();
                Helpers.httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", TeamsMigrate.Utils.Auth.AccessToken);
                Helpers.httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                log.Debug("POST " + O365.MsGraphBetaEndpoint + "teams/" + selectedTeamId + "/channels/" + channelId + "/completeMigration");

                if (Program.CmdOptions.ReadOnly)
                {
                    log.Debug("skip operation due to readonly mode");
                }

                var completeMigrationResponseMessage = Helpers.httpClient.PostAsync(O365.MsGraphBetaEndpoint + "teams/" + selectedTeamId + "/channels/" + channelId + "/completeMigration", new StringContent("", Encoding.UTF8, "application/json")).Result;

                if (!completeMigrationResponseMessage.IsSuccessStatusCode)
                {
                    log.Error("Failed to complete channel migration");
                    log.Debug(completeMigrationResponseMessage.Content.ReadAsStringAsync().Result);
                }
            }
            catch (Exception ex)
            {
                log.Error("Failed to complete channel migration");
                log.Debug("Failure", ex);
            }
        }

        public static List<MsTeams.Channel> GetExistingChannelsInMsTeams(string teamId)
        {
            MsTeams.Team msTeamsTeam = new MsTeams.Team();

            Helpers.httpClient.DefaultRequestHeaders.Clear();
            Helpers.httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", TeamsMigrate.Utils.Auth.AccessToken);
            Helpers.httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            var url = O365.MsGraphBetaEndpoint + "teams/" + teamId + "/channels";
            log.Debug("GET " + url);
            var httpResponseMessage = Helpers.httpClient.GetAsync(url).Result;
            log.Debug(httpResponseMessage);
            if (httpResponseMessage.IsSuccessStatusCode)
            {
                var httpResultString = httpResponseMessage.Content.ReadAsStringAsync().Result;
                log.Debug(httpResultString);
                msTeamsTeam = JsonConvert.DeserializeObject<MsTeams.Team>(httpResultString);
            }

            return msTeamsTeam.value;
        }

        internal static void AssignChannelsMembership(string selectedTeamId, List<Combined.ChannelsMapping> msTeamsChannelsWithSlackProps, List<ViewModels.SimpleUser> slackUserList)
        {
            var teamUsers = new HashSet<string>();
            foreach (var channel in msTeamsChannelsWithSlackProps)
            {
                var existingMsTeams = GetExistingChannelByName(channel.displayName);
                if (existingMsTeams == null)
                {
                    continue;
                }
                int i = 1;
                using (var progress = new ProgressBar(String.Format("Update '{0}' membership", channel.displayName)))
                {
                    foreach (var member in channel.members)
                    {
                        progress.Report((double)i++ / channel.members.Count);
                        if (String.IsNullOrEmpty(member))
                        {
                            continue;
                        }
                        var user = slackUserList.FirstOrDefault(u => member.Equals(u.userId));
                        if (user != null)
                        {
                            var userId = Users.GetUserIdByName(user.name);
                            if (String.IsNullOrEmpty(userId))
                            {
                                log.DebugFormat("Missing user {0}", user.name + "@" + Program.CmdOptions.Domain);
                                continue;
                            }
                            if (!teamUsers.Contains(member))
                            {
                                log.DebugFormat("Add {0} to team {1}", user.name, selectedTeamId);
                                if (Users.AddMemberTeam(selectedTeamId, userId))
                                {
                                    teamUsers.Add(member);
                                }
                            }
                            if (!channel.membershipType.Equals("standard"))
                            {
                                log.DebugFormat("Add {0} to channel {1}", user.name, channel.id);
                                Users.AddMemberChannel(selectedTeamId, channel.id, userId);
                            }
                        }
                        else
                        {
                            log.DebugFormat("Missing member {0}", member);
                        }
                    }
                }
            }
        }

        internal static void AssignTeamOwnerships(string selectedTeamId)
        {
            Console.Write("Do you want to assign ownership? (y|n): ");
            var completeMigration = Console.ReadLine();
            if (completeMigration.StartsWith("y", StringComparison.CurrentCultureIgnoreCase))
            {
                if (String.IsNullOrEmpty(TeamsMigrate.Utils.Auth.UserToken))
                {
                    TeamsMigrate.Utils.Auth.UserLogin();
                }

                Helpers.httpClient.DefaultRequestHeaders.Clear();
                Helpers.httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", TeamsMigrate.Utils.Auth.UserToken);
                Helpers.httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                Users.AddOwner(selectedTeamId, O365.getUserGuid(TeamsMigrate.Utils.Auth.UserToken, "me"));
            }
        }
    }
}
