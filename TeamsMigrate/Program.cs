using System;
using System.IO;
using CommandLine;
using log4net;
using log4net.Core;
using System.Collections.Generic;

namespace TeamsMigrate
{
    class Program
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(Program));

        private static System.Timers.Timer aTimer;

        public static Options CmdOptions { get; internal set; }

        public const string aadResourceAppId = "00000003-0000-0000-c000-000000000000";

        public static bool SkipCleanup = false;

        static void Main(string[] args)
        {
            log4net.Config.XmlConfigurator.Configure();
            LogManager.GetRepository().Threshold = Level.Info;

            Parser.Default.ParseArguments<Options>(args)
                   .WithParsed(options => { CmdOptions = options; });

            if (CmdOptions.Verbose)
            {
                LogManager.GetRepository().Threshold = Level.Debug;
            }

            string slackArchiveBasePath = "";
            string slackArchiveTempPath = "";
            string channelsPath = "";

            Console.CancelKeyPress += delegate (object sender, ConsoleCancelEventArgs e)
            {
                try
                {
                    Utils.Files.CleanUpTempDirectoriesAndFiles(slackArchiveTempPath);
                }
                catch (Exception ex)
                {
                    log.Error("Failed to complete cleanup");
                    log.Debug("Failure", ex);
                }
            };

            log.Info("Tenant is " + CmdOptions.TenantId);
            log.Info("Application ID is " + CmdOptions.ClientId);
            log.Info("Redirect URI is " + CmdOptions.AadRedirectUri);
            log.InfoFormat("Tenant admin consent URL is https://login.microsoftonline.com/common/oauth2/authorize?response_type=id_token&client_id={0}&redirect_uri={1}&prompt=admin_consent&nonce={2}", CmdOptions.ClientId, CmdOptions.AadRedirectUri, Guid.NewGuid().ToString());

            TeamsMigrate.Utils.Auth.AccessToken = TeamsMigrate.Utils.Auth.Login();

            if (String.IsNullOrEmpty(TeamsMigrate.Utils.Auth.AccessToken))
            {
                log.Info("Something went wrong.  Please try again!");
                Environment.Exit(1);
            }
            else
            {
                log.Info("You've successfully signed in.");
            }

            aTimer = new System.Timers.Timer();
            aTimer.Interval = 30 * 60 * 1000;

            aTimer.Elapsed += (Object source, System.Timers.ElapsedEventArgs e) =>
            {
                TeamsMigrate.Utils.Auth.AccessToken = TeamsMigrate.Utils.Auth.Login();
            };

            aTimer.AutoReset = true;
            aTimer.Enabled = true;
            string teamName;

            if (!CmdOptions.ExportPath.EndsWith(".zip", StringComparison.CurrentCulture))
            {
                if (!Directory.Exists(CmdOptions.ExportPath))
                {
                    log.ErrorFormat("Directory {0} does not exist. Exit.", CmdOptions.ExportPath);
                    Environment.Exit(0);
                }
                teamName = Path.GetDirectoryName(CmdOptions.ExportPath);
                slackArchiveBasePath = CmdOptions.ExportPath;
                slackArchiveTempPath = CmdOptions.ExportPath;
                channelsPath = Path.Combine(slackArchiveBasePath, "channels.json");
                SkipCleanup = true;
            }
            else
            {
                slackArchiveTempPath = Path.GetTempFileName();
                slackArchiveBasePath = Utils.Files.DecompressSlackArchiveFile(CmdOptions.ExportPath, slackArchiveTempPath);
                teamName = CmdOptions.ExportPath.ToString().Split(".")[0];
                channelsPath = Path.Combine(slackArchiveBasePath, "channels.json");
            }

            log.DebugFormat("Use directory {0} ({1})", slackArchiveBasePath, slackArchiveTempPath);

            var slackChannelsToMigrate = Utils.Channels.ScanSlackChannelsJson(channelsPath);

            if (File.Exists(Path.Combine(slackArchiveBasePath, "groups.json")))
            {
                slackChannelsToMigrate.AddRange(Utils.Channels.ScanSlackChannelsJson(Path.Combine(slackArchiveBasePath, "groups.json"), "private"));
            }

            var slackUserList = Utils.Users.ScanUsers(Path.Combine(slackArchiveBasePath, "users.json"));

            // Ask for the existing team ID
            Console.Write("Enter the existing Team ID: ");
            string selectedTeamId = Console.ReadLine();

            // Ask for channel mappings
            Dictionary<string, string> channelMappings = new Dictionary<string, string>();
            foreach (var slackChannel in slackChannelsToMigrate)
            {
                Console.Write($"Enter the corresponding MS Teams channel ID for Slack channel '{slackChannel.channelName}': ");
                string msTeamsChannelId = Console.ReadLine();
                channelMappings.Add(slackChannel.channelId, msTeamsChannelId);
            }

            if (CmdOptions.MigrateMessages)
            {
                Utils.Messages.ScanMessagesByChannel(slackChannelsToMigrate.Select(sc => new Models.Combined.ChannelsMapping
                {
                    slackChannelId = sc.channelId,
                    slackChannelName = sc.channelName,
                    displayName = sc.channelName,
                    id = channelMappings[sc.channelId]
                }).ToList(), slackArchiveTempPath, slackUserList, selectedTeamId, CmdOptions.MigrateFiles);
            }

            if (!Program.CmdOptions.ReadOnly)
            {
                // Assign ownerships and memberships
                Utils.Channels.AssignTeamOwnerships(selectedTeamId);
                Utils.Channels.AssignChannelsMembership(selectedTeamId, slackChannelsToMigrate.Select(sc => new Models.Combined.ChannelsMapping
                {
                    slackChannelId = sc.channelId,
                    slackChannelName = sc.channelName,
                    displayName = sc.channelName,
                    id = channelMappings[sc.channelId]
                }).ToList(), slackUserList);

                TeamsMigrate.Utils.Users.UsersCleanup();
            }

            Utils.Files.CleanUpTempDirectoriesAndFiles(slackArchiveTempPath);
        }
    }
}
