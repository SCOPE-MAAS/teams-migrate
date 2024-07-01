using System;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Threading;
using System.Net.Http;
using TeamsMigrate.Models;

namespace TeamsMigrate.Utils
{
    public class FileAttachments
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(FileAttachments));

        static int index = 1;
        public static async Task ArchiveMessageFileAttachments(string selectedTeamId, List<Combined.AttachmentsMapping> combinedAttachmentsMapping, string channelSubFolder, int maxDls = 10)
        {
            var tasks = new List<Task>();
            index = 1;
            using (var progress = new ProgressBar("Uploading files"))
            {
                SemaphoreSlim semaphore = new SemaphoreSlim(maxDls);

                foreach (var v in combinedAttachmentsMapping)
                {
                    await semaphore.WaitAsync();
                    tasks.Add(GetAndUploadFileToTeamsChannel(selectedTeamId, semaphore, v, channelSubFolder, progress, combinedAttachmentsMapping.Count));
                }

                await Task.WhenAll(tasks);
            }
        }

        static async Task GetAndUploadFileToTeamsChannel(string selectedTeamId, SemaphoreSlim semaphore, Combined.AttachmentsMapping combinedAttachmentsMapping, string channelSubFolder, ProgressBar progress, int count)
        {
            Tuple<string, string> fileIdAndUrl;
            try
            {
                if (Program.CmdOptions.ReadOnly)
                {
                    log.Debug("skip operation due to readonly mode");
                }
                else
                {
                    string fileToUpload = "";
                    if (!combinedAttachmentsMapping.attachmentUrl.StartsWith("/"))
                    {
                        log.Debug("Downloading attachment to local file system " + combinedAttachmentsMapping.attachmentId);
                        var request = new HttpClient();
                        using (HttpResponseMessage response = await request.GetAsync(combinedAttachmentsMapping.attachmentUrl, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(false))
                        {
                            using (Stream streamToReadFrom = await response.Content.ReadAsStreamAsync())
                            {
                                fileToUpload = Path.GetTempFileName();
                                using (Stream streamToWriteTo = File.Open(fileToUpload, FileMode.Create))
                                {
                                    await streamToReadFrom.CopyToAsync(streamToWriteTo);
                                }
                            }
                        }
                    }
                    else
                    {
                        fileToUpload = combinedAttachmentsMapping.attachmentUrl;
                    }
                    var pathToItem = "/" + combinedAttachmentsMapping.msChannelName + "/slack-files/" + combinedAttachmentsMapping.attachmentId + "/" + combinedAttachmentsMapping.attachmentFileName;
                    fileIdAndUrl = await UploadFileToTeamsChannel(selectedTeamId, fileToUpload, pathToItem);
                    combinedAttachmentsMapping.msSpoId = fileIdAndUrl.Item1;
                    combinedAttachmentsMapping.msSpoUrl = fileIdAndUrl.Item2;
                    File.Delete(fileToUpload);
                    log.Debug("Deleting local copy of attachment " + combinedAttachmentsMapping.attachmentId);
                }
            }
            catch (Exception ex)
            {
                log.Debug("Failed to upload file " + combinedAttachmentsMapping.attachmentId + " " + combinedAttachmentsMapping.attachmentUrl);
                log.Debug("Failure", ex);
            }
            finally
            {
                progress.Report((double)index++ / count);
                semaphore.Release();
            }
            return;
        }

        public static Tuple<string, string> CheckIfFileExistsOnTeamsChannel(string selectedTeamId, string pathToItem)
        {
            var authHelper = new Utils.O365.AuthenticationHelper() { AccessToken = TeamsMigrate.Utils.Auth.AccessToken };
            Microsoft.Graph.GraphServiceClient gcs = new Microsoft.Graph.GraphServiceClient(authHelper);

            Microsoft.Graph.DriveItem fileExistsResult = null;
            try
            {
                fileExistsResult = gcs.Groups[selectedTeamId].Drive.Root.ItemWithPath(pathToItem).Request().GetAsync().Result;
            }
            catch
            {
                fileExistsResult = null;
            }

            if (fileExistsResult == null)
            {
                return new Tuple<string, string>("", "");
            }
            log.Debug("Attachment already exists.  We won't replace it. " + pathToItem);
            return new Tuple<string, string>(fileExistsResult.Id, fileExistsResult.WebUrl);
        }

        public static async Task<Tuple<string, string>> UploadFileToTeamsChannel(string selectedTeamId, string filePath, string pathToItem)
        {
            var authHelper = new Utils.O365.AuthenticationHelper() { AccessToken = TeamsMigrate.Utils.Auth.AccessToken };
            Microsoft.Graph.GraphServiceClient gcs = new Microsoft.Graph.GraphServiceClient(authHelper);

            var fileExists = CheckIfFileExistsOnTeamsChannel(selectedTeamId, pathToItem);

            if (Program.CmdOptions.ReadOnly)
            {
                log.Debug("skip operation due to readonly mode");
                return new Tuple<string, string>(fileExists.Item1, fileExists.Item2);
            }

            if (fileExists.Item1 != "")
            {
                return new Tuple<string, string>(fileExists.Item1, fileExists.Item2);
            }

            Microsoft.Graph.UploadSession uploadSession = await gcs.Groups[selectedTeamId].Drive.Root.ItemWithPath(pathToItem).CreateUploadSession().Request().PostAsync();

            try
            {
                log.DebugFormat("Trying to upload file {0} ({1}) ", pathToItem, filePath);
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    var maxChunkSize = 320 * 1024; // 320 KB

                    var chunkedUploadProvider = new Microsoft.Graph.ChunkedUploadProvider(uploadSession, gcs, fs, maxChunkSize);

                    var chunkRequests = chunkedUploadProvider.GetUploadChunkRequests();
                    var trackedExceptions = new List<Exception>();

                    Microsoft.Graph.DriveItem itemResult = null;

                    foreach (var request in chunkRequests)
                    {
                        var result = await chunkedUploadProvider.GetChunkRequestResponseAsync(request, trackedExceptions);

                        if (result.UploadSucceeded)
                        {
                            itemResult = result.ItemResponse;
                        }
                    }
                    log.Debug("Upload of attachment to MS Teams completed " + pathToItem);
                    log.Debug("SPo ID is " + itemResult.Id + " URL: " + itemResult.WebUrl);
                    return new Tuple<string, string>(itemResult.Id, itemResult.WebUrl);
                }
            }
            catch (Exception ex)
            {
                log.Error("Attachment could not be uploaded");
                log.Debug("Failure", ex);
            }

            return new Tuple<string, string>("", "");
        }
    }
}
