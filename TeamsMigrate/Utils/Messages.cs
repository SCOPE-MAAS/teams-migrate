using TeamsMigrate.Models;
using TeamsMigrate.ViewModels;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace TeamsMigrate.Utils
{
    public class Messages
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(Messages));

        private static Dictionary<string, string> SlackToTeamsIdsMapping = new Dictionary<string, string>();

        public static void ScanMessagesByChannel(List<Models.Combined.ChannelsMapping> channelsMapping, string basePath,
            List<ViewModels.SimpleUser> slackUserList, String selectedTeamId, bool copyFileAttachments)
        {
            int i = 1;
            foreach (var channel in channelsMapping)
            {
                List<Models.Combined.AttachmentsMapping> channelAttachmentsToUpload = null;
                try
                {
                    log.InfoFormat("Migrating messages in channel {0} ({1} out of {2})", channel.slackChannelName, i++, channelsMapping.Count);
                    channelAttachmentsToUpload = GetAndUploadMessages(channel, basePath, slackUserList, selectedTeamId, copyFileAttachments);
                }
                catch (Exception ex)
                {
                    log.Error("Failed to upload messages");
                    log.Debug("Failure", ex);
                }
            }
        }

        static List<Models.Combined.AttachmentsMapping> GetAndUploadMessages(Models.Combined.ChannelsMapping channelsMapping, string basePath,
            List<ViewModels.SimpleUser> slackUserList, String selectedTeamId, bool copyFileAttachments)
        {
            var messageList = new List<ViewModels.SimpleMessage>();
            messageList.Clear();

            var messageListJsonSource = new JArray();
            messageListJsonSource.Clear();

            List<Models.Combined.AttachmentsMapping> attachmentsToUpload = new List<Models.Combined.AttachmentsMapping>();
            attachmentsToUpload.Clear();

            var files = Directory.GetFiles(Path.Combine(basePath, channelsMapping.slackChannelName));
            using (var progress = new ProgressBar("Loading message files"))
            {
                int i = 1;
                foreach (var file in files)
                {
                    progress.Report((double)i++ / files.Length);
                    log.Debug("File " + file);
                    using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
                    using (StreamReader sr = new StreamReader(fs))
                    using (JsonTextReader reader = new JsonTextReader(sr))
                    {
                        while (reader.Read())
                        {
                            if (reader.TokenType == JsonToken.StartObject)
                            {
                                JObject obj = JObject.Load(reader);

                                var messageTs = obj.SelectToken("ts").ToString();
                                var messageText = (string)obj.SelectToken("text");
                                if (messageText.EndsWith("has joined the channel") || messageText.EndsWith("has joined the group"))
                                    continue;

                                var messageId = (string)obj.SelectToken("ts");
                                string rootId = (string)obj.SelectToken("thread_ts");

                                var messageSender = FindMessageSender(obj, slackUserList);
                                if (messageSender == null)
                                {
                                    log.Warn("Original sender not found. Skipping message.");
                                    continue;
                                }

                                List<ViewModels.SimpleMessage.FileAttachment> fileAttachments = new List<ViewModels.SimpleMessage.FileAttachment>();
                                ViewModels.SimpleMessage.FileAttachment fileAttachment = null;
                                var filesObject = (JArray)obj.SelectToken("files");
                                if (filesObject != null)
                                {
                                    foreach (var f in filesObject)
                                    {
                                        var fileUrl = (string)f.SelectToken("url_private");
                                        var fileId = (string)f.SelectToken("id");
                                        var fileMode = (string)f.SelectToken("mode");
                                        var fileName = (string)f.SelectToken("name");

                                        if (fileMode != "external" && fileId != null && fileUrl != null)
                                        {
                                            log.Debug("Message attachment found with ID " + fileId);
                                            attachmentsToUpload.Add(new Models.Combined.AttachmentsMapping
                                            {
                                                attachmentId = fileId,
                                                attachmentUrl = fileUrl,
                                                attachmentChannelId = channelsMapping.slackChannelId,
                                                attachmentFileName = fileName,
                                                msChannelName = channelsMapping.displayName
                                            });

                                            fileAttachment = new ViewModels.SimpleMessage.FileAttachment
                                            {
                                                id = fileId,
                                                originalName = (string)f.SelectToken("name"),
                                                originalTitle = (string)f.SelectToken("title"),
                                                originalUrl = (string)f.SelectToken("permalink")
                                            };
                                            fileAttachments.Add(fileAttachment);
                                        }
                                    }
                                }

                                List<ViewModels.SimpleMessage.Attachments> attachmentsList = new List<ViewModels.SimpleMessage.Attachments>();
                                List<ViewModels.SimpleMessage.Attachments.Fields> fieldsList = new List<ViewModels.SimpleMessage.Attachments.Fields>();
                                var attachmentsObject = (JArray)obj.SelectToken("attachments");
                                if (attachmentsObject != null)
                                {
                                    foreach (var attachmentItem in attachmentsObject)
                                    {
                                        if (String.IsNullOrEmpty(rootId))
                                        {
                                            rootId = (string)attachmentItem.SelectToken("ts");
                                        }
                                        var attachmentText = (string)attachmentItem.SelectToken("text");
                                        var attachmentTextFallback = (string)attachmentItem.SelectToken("fallback");

                                        var attachmentItemToAdd = new ViewModels.SimpleMessage.Attachments();

                                        if (!String.IsNullOrEmpty(attachmentText))
                                        {
                                            attachmentItemToAdd.text = attachmentText;
                                        }
                                        else if (!String.IsNullOrEmpty(attachmentTextFallback))
                                        {
                                            attachmentItemToAdd.text = attachmentTextFallback;
                                        }

                                        var attachmentServiceName = (string)attachmentItem.SelectToken("service_name");
                                        if (!String.IsNullOrEmpty(attachmentServiceName))
                                        {
                                            attachmentItemToAdd.service_name = attachmentServiceName;
                                        }

                                        var attachmentFromUrl = (string)attachmentItem.SelectToken("from_url");
                                        if (!String.IsNullOrEmpty(attachmentFromUrl))
                                        {
                                            attachmentItemToAdd.url = attachmentFromUrl;
                                        }

                                        var attachmentColor = (string)attachmentItem.SelectToken("color");
                                        if (!String.IsNullOrEmpty(attachmentColor))
                                        {
                                            attachmentItemToAdd.color = attachmentColor;
                                        }

                                        var fieldsObject = (JArray)attachmentItem.SelectToken("fields");
                                        if (fieldsObject != null)
                                        {
                                            fieldsList.Clear();
                                            foreach (var fieldItem in fieldsObject)
                                            {
                                                fieldsList.Add(new ViewModels.SimpleMessage.Attachments.Fields()
                                                {
                                                    title = (string)fieldItem.SelectToken("title"),
                                                    value = (string)fieldItem.SelectToken("value"),
                                                    shortWidth = (bool)fieldItem.SelectToken("short")
                                                });
                                            }
                                            attachmentItemToAdd.fields = fieldsList;
                                        }
                                        else
                                        {
                                            attachmentItemToAdd.fields = null;
                                        }
                                        attachmentsList.Add(attachmentItemToAdd);
                                    }
                                }
                                else
                                {
                                    attachmentsList = null;
                                }

                                messageList.Add(new ViewModels.SimpleMessage
                                {
                                    id = messageId,
                                    text = HandleContent(messageText, slackUserList),
                                    ts = UnixTimeStampToDateTime(Convert.ToDouble(messageTs)).ToUniversalTime()
                                        .ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'ffffff'Z'"),
                                    user = messageSender,
                                    userId = Users.GetOrCreateId(messageSender, slackUserList, Program.CmdOptions.Domain),
                                    fileAttachments = fileAttachments,
                                    attachments = attachmentsList,
                                    rootId = rootId,
                                });
                            }
                        }
                    }
                }
            }
            if (copyFileAttachments)
            {
                Utils.FileAttachments.ArchiveMessageFileAttachments(selectedTeamId, attachmentsToUpload, "fileattachments").Wait();

                foreach (var messageItem in messageList)
                {
                    foreach (ViewModels.SimpleMessage.FileAttachment attachment in messageItem.fileAttachments)
                    {
                        if (attachment != null)
                        {
                            var messageItemWithFileAttachment = attachmentsToUpload.Find(w => String.Equals(attachment.id, w.attachmentId, StringComparison.CurrentCultureIgnoreCase));
                            if (messageItemWithFileAttachment != null)
                            {
                                attachment.spoId = messageItemWithFileAttachment.msSpoId;
                                attachment.spoUrl = messageItemWithFileAttachment.msSpoUrl;
                            }
                        }
                    }
                }
            }

            Utils.Messages.ImportMessages(basePath, channelsMapping, messageList, selectedTeamId);

            return attachmentsToUpload;
        }

        private static string HandleContent(string messageText, List<ViewModels.SimpleUser> slackUserList)
        {
            messageText = messageText.Replace("\n", "<br/>\n");

            Regex reg = new Regex(@"<@\w{5,20}>", RegexOptions.IgnoreCase);
            Match match;

            List<string> results = new List<string>();
            for (match = reg.Match(messageText); match.Success; match = match.NextMatch())
            {
                if (!(results.Contains(match.Value)))
                    results.Add(match.Value);
            }

            foreach (var item in results)
            {
                try
                {
                    var user = slackUserList.First(u => item.Replace("<@", "").Replace(">", "").Equals(u.userId));
                    messageText = messageText.Replace(item.ToString(), "@" + user.real_name.ToString());
                }
                catch
                {
                    log.DebugFormat("Error replacing user '{0}'", item);
                }
            }

            reg = new Regex(@"<(?<Protocol>\w+):\/\/(?<Domain>[\w@][\w.:@]+)\/?[\w\.?=%&=\-@/$|,]*>");
            match = reg.Match(messageText);
            for (match = reg.Match(messageText); match.Success; match = match.NextMatch())
            {
                var url = match.Value.TrimEnd('>').TrimStart('<');
                string[] href = url.Split("|");
                messageText = messageText.Replace(match.Value, String.Format("<a href='{0}'>{1}</a>", href.Length > 0 ? href[0] : url, href.Length > 1 ? href[1] : url));
                log.DebugFormat("Found URL in message {0}", messageText);
            }

            return messageText;
        }

        private static void ImportMessages(string basePath, Combined.ChannelsMapping channelsMapping, List<ViewModels.SimpleMessage> messageList, string selectedTeamId)
        {
            int i = 1;
            using (var progress = new ProgressBar("Importing messages"))
            {
                foreach (ViewModels.SimpleMessage message in messageList)
                {
                    progress.Report((double)i++ / messageList.Count);
                    ImportExternalMessage(selectedTeamId, channelsMapping.id, message);
                }
            }
        }

        internal static void ImportExternalMessage(string teamId, string channelId, SimpleMessage message, bool retry = true)
        {
            if (string.IsNullOrEmpty(message.text))
            {
                log.Debug("Empty message: " + JsonConvert.SerializeObject(message));
                return;
            }

            Helpers.httpClient.DefaultRequestHeaders.Clear();
            Helpers.httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", TeamsMigrate.Utils.Auth.AccessToken);
            Helpers.httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var converted = new ExternalMessage(message);
            var createTeamsChannelPostData = JsonConvert.SerializeObject(converted);

            string url;
            if (String.IsNullOrEmpty(message.rootId))
            {
                url = O365.MsGraphBetaEndpoint + "teams/" + teamId + "/channels/" + channelId + "/messages";
            }
            else
            {
                if (SlackToTeamsIdsMapping.ContainsKey(message.rootId))
                    url = O365.MsGraphBetaEndpoint + "teams/" + teamId + "/channels/" + channelId + "/messages/" + SlackToTeamsIdsMapping[message.rootId] + "/replies";
                else
                {
                    log.Debug("Missing key: " + message.rootId);
                    url = O365.MsGraphBetaEndpoint + "teams/" + teamId + "/channels/" + channelId + "/messages";
                }
            }

            log.Debug("POST " + url);
            log.Debug(createTeamsChannelPostData);

            if (Program.CmdOptions.ReadOnly)
            {
                log.Debug("skip operation due to readonly mode");
                return;
            }

            try
            {
                HttpResponseMessage httpResponseMessage = Helpers.httpClient.PostAsync(url, new StringContent(createTeamsChannelPostData, Encoding.UTF8, "application/json")).Result;
                log.Debug(httpResponseMessage.Content.ReadAsStringAsync().Result);
                if (!httpResponseMessage.IsSuccessStatusCode)
                {
                    log.Debug("Orig message: " + JsonConvert.SerializeObject(message));
                    if (httpResponseMessage.Content.ReadAsStringAsync().Result.Contains("TooManyRequests") && retry)
                    {
                        Thread.Sleep(2000);
                        ImportExternalMessage(teamId, channelId, message, false);
                    }
                    return;
                }
                dynamic response = JObject.Parse(httpResponseMessage.Content.ReadAsStringAsync().Result);
                if (response != null && !String.IsNullOrEmpty(message.id))
                {
                    string messageId = response.id;
                    log.DebugFormat("Set message ID {0},{1}", message.id, messageId);
                    SlackToTeamsIdsMapping.Add(message.id, messageId);
                }
            }
            catch (Exception e)
            {
                log.Debug("Failed to import message ", e);
                log.Debug("Orig message: " + JsonConvert.SerializeObject(message));
            }
        }

        static string FindMessageSender(JObject obj, List<ViewModels.SimpleUser> slackUserList)
        {
            var user = (string)obj.SelectToken("user");
            if (!String.IsNullOrEmpty(user))
            {
                if (user != "USLACKBOT")
                {
                    var simpleUser = slackUserList.FirstOrDefault(w => w.userId == user);
                    if (simpleUser != null)
                    {
                        return simpleUser.name;
                    }
                }
                else
                {
                    return "SlackBot";
                }
            }
            else if (!(String.IsNullOrEmpty((string)obj.SelectToken("username"))))
            {
                return (string)obj.SelectToken("username");
            }
            else if (!(String.IsNullOrEmpty((string)obj.SelectToken("bot_id"))))
            {
                return (string)obj.SelectToken("bot_id");
            }

            log.Warn("Original sender not found. Skipping message.");
            return null;
        }
    }
}
