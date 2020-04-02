using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace eni
{
    public static class blob2spo
    {
        private static HttpClient client;

        static blob2spo()
        {
            client = new HttpClient();
            client.Timeout = TimeSpan.FromMinutes(60);
        }
        
        [FunctionName("blob2spo")]
        public static async Task Run(
            [BlobTrigger("drop/{name}", Connection = "ricchiblob_STORAGE")]Stream myBlob,
            string name,
            IDictionary<string, string> metaData, // See https://dontcodetired.com/blog/post/Getting-Blob-Metadata-When-Using-Azure-Functions-Blob-Storage-Triggers
            ILogger log)
        {
            log.LogInformation($"C# Blob trigger function Processed blob\n Name:{name} \n Size: {myBlob.Length} Bytes");

            log.LogInformation("Getting token");
            var jwt = await GetJwt(log);
            log.LogDebug($"Got {jwt}");

            log.LogInformation("Uploading file");
            await UploadFile(myBlob, name, jwt, log);
        }

        public static string GetSetting(string name)
        {
            return name + ": " +
                System.Environment.GetEnvironmentVariable(name, EnvironmentVariableTarget.Process);
        }

        private static async Task UploadFile(Stream myBlob, string name, string token, ILogger log)
        {
            try
            {
                string tenantUrl = GetSetting("TENANT_URL");
                string siteUrl = GetSetting("SITE_RELATIVE_URL");
                string folderPath = GetSetting("SITE_FOLDER_PATH");

                HttpRequestMessage req = new HttpRequestMessage();
                req.Content = new StreamContent(myBlob);
                req.RequestUri = new Uri($"{tenantUrl}{siteUrl}/_api/web/GetFolderByServerRelativeUrl('{siteUrl}{folderPath}')/Files/add(url='{name}',overwrite=true)");
                req.Method = HttpMethod.Post;
                req.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                var response = await client.SendAsync(req);

                if (!response.IsSuccessStatusCode)
                {
                    var text = await response.Content.ReadAsStringAsync();
                    throw new ApplicationException(text);
                }
            }
            catch (System.Exception ex)
            {
                log.LogError(ex, "Error uploading");
                throw;
            }
        }

        private static async Task<string> GetJwt(ILogger log)
        {
            string clientId = GetSetting("SP_APP_ID");
            string clientSecret = GetSetting("SP_APP_SECRET");
            string aadTenantId = GetSetting("AAD_TENANT_ID");
            var body = new FormUrlEncodedContent(new Dictionary<string, string> {
                { "grant_type", "client_credentials" },
                { "client_secret", clientSecret },
                { "resource", $"00000003-0000-0ff1-ce00-000000000000/enispa.sharepoint.com@{aadTenantId}" },
                { "client_id", $"{clientId}@{aadTenantId}" }
            });
            var result = await client.PostAsync(
                $"https://accounts.accesscontrol.windows.net/{aadTenantId}/tokens/OAuth/2",
                body);

            result.EnsureSuccessStatusCode();

            string json = await result.Content.ReadAsStringAsync();
            log.LogDebug($"Got json {json}");
            dynamic value = JsonConvert.DeserializeObject<dynamic>(json);

            return value.access_token;
        }
    }
}
