using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using System.Net;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Company.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph.Models;
using Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession;
using Microsoft.Graph;
using Microsoft.Graph.Models.ODataErrors;

namespace Company.Function
{
    public class ConvertApi
    {
        private readonly ILogger _logger;
        private readonly IGraphClientService _graphClientService;
        private readonly IConfiguration _config;

        public ConvertApi(ILoggerFactory loggerFactory, IGraphClientService graphClientService, IConfiguration config)
        {
            _logger = loggerFactory.CreateLogger<ConvertApi>();
            _graphClientService = graphClientService;
            _config = config;
        }

        public Task<Stream> DownloadDoc(string driveId, string fileId, GraphServiceClient graphClient){
            //var requestInfo = graphClient.Drives[driveId].Items[fileId].Content.ToGetRequestInformation();
            try {
                _logger.LogInformation("Downloading");
                return graphClient.Drives[driveId].Items[fileId].Content
                .GetAsync((requestConfiguration) =>
                {
                    requestConfiguration.QueryParameters.Format = "pdf";
                });
            } 
            catch (ODataError odataError)
            {
                Console.WriteLine(odataError.Error?.Code);
                Console.WriteLine(odataError.Error?.Message);
                throw;
            } 
        }

        [Function("SmallDocConversion")]
        public async Task<IActionResult> RunSmallDocConversionAsync([HttpTrigger(AuthorizationLevel.Anonymous, "post")] HttpRequestData req)
        {
            _logger.LogInformation("C# HTTP trigger function processed a request.");

            var graphClient = _graphClientService.GetAppGraphClient();
            var driveId = _config["DRIVE_ID"];
            var itemId = _config["ITEM_ID"];
            // var fileName = req.Headers.GetValues("file-name").FirstOrDefault("uploaded-file");

            string contentType = req.Headers.GetValues("Content-Type").FirstOrDefault("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            string fileName = $"{Guid.NewGuid()}{MimeTypes.MimeTypeMap.GetExtension(contentType)}";
            _logger.LogInformation("Filename: " + fileName);

            var uploadDriveItem = await graphClient.Drives[driveId].Items[itemId].Children[fileName].Content.PutAsync(req.Body, requestConfig => {
                requestConfig.Headers.Add("Content-Type", req.Headers.GetValues("Content-Type"));
            });

            var uploadedFileId = uploadDriveItem.Id;
            _logger.LogInformation(uploadedFileId);

            var response = req.CreateResponse(HttpStatusCode.OK);
            var filecontent = await DownloadDoc(driveId, uploadedFileId, graphClient);
            return new FileStreamResult(filecontent, "application/pdf");
        }

        // public async Task<HttpResponseData> DownloadDocV2(string driveId, string fileId, HttpRequestData req, GraphServiceClient graphClient){
        //     _logger.LogInformation("Downloading Doc");
        //     var requestInfo = graphClient.Drives[driveId].Items[fileId].Content.ToGetRequestInformation();
        //     requestInfo.UrlTemplate += "{?format}";
        //     requestInfo.QueryParameters.Add("format", "pdf");
        //     var stream = await graphClient.RequestAdapter.SendPrimitiveAsync<Stream>(requestInfo);
        //     var response = req.CreateResponse(HttpStatusCode.OK);
        //     response.Headers.Add("Content-Type", "application/octet-stream");
        //     response.Body = stream;
        //     return response;
        // }
        public async Task<string> UploadLargeDoc(string driveId, string itemId, string fileName, Stream body, GraphServiceClient graphClient){
            _logger.LogInformation("Uploading Large Doc.");
            var uploadSessionRequestBody = new CreateUploadSessionPostRequestBody
            {
                Item = new DriveItemUploadableProperties
                {
                    AdditionalData = new Dictionary<string, object>
                    {
                        { "@microsoft.graph.conflictBehavior", "replace" },
                    },
                },
            };

            // Create the upload session
            var uploadSession = await graphClient.Drives[driveId].Items[itemId].ItemWithPath(fileName).CreateUploadSession.PostAsync(uploadSessionRequestBody);

            // Max slice size must be a multiple of 320 KiB
            int maxSliceSize = 320 * 1024;
            var fileUploadTask = new LargeFileUploadTask<DriveItem>(
                uploadSession, body, maxSliceSize, graphClient.RequestAdapter);

            var totalLength = body.Length;
            // Create a callback that is invoked after each slice is uploaded
            IProgress<long> progress = new Progress<long>(prog =>
            {
                _logger.LogInformation($"Uploaded {prog} bytes of {totalLength} bytes");
            });

            var uploadResult = await fileUploadTask.UploadAsync(progress);

            _logger.LogInformation(uploadResult.UploadSucceeded ?
                $"Upload complete, item ID: {uploadResult.ItemResponse.Id}" :
                "Upload failed");

            return uploadResult.ItemResponse.Id;
        }

        [Function("LargeDocConversion")]
        public async Task<IActionResult> RunLargeDocConversionAsync([HttpTrigger(AuthorizationLevel.Anonymous, "post")] HttpRequest req)
        {
            _logger.LogInformation("C# HTTP trigger function processed a request.");

            var graphClient = _graphClientService.GetAppGraphClient();
            var driveId = _config["DRIVE_ID"];
            var itemId = _config["ITEM_ID"];
            string contentType = req.Headers["Content-Type"];
            //string contentType = req.Headers.GetValues("Content-Type").FirstOrDefault("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            string fileName = $"{Guid.NewGuid()}{MimeTypes.MimeTypeMap.GetExtension(contentType)}";

            _logger.LogInformation("filename: " + fileName);
            var stream = new MemoryStream();
            await req.Body.CopyToAsync(stream);

            var fileId = await UploadLargeDoc(driveId, itemId, fileName, stream, graphClient);

            var filecontent = await DownloadDoc(driveId, fileId, graphClient);
            return new FileStreamResult(filecontent, "application/pdf");
        }
    }
}
