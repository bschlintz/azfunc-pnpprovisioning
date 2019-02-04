using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using System.Security.Cryptography.X509Certificates;
using Microsoft.SharePoint.Client;
using System.Threading;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using System;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using Newtonsoft.Json;

namespace PnPProvisioningFunction
{
    public static class ExtractAndApply
    {
        static readonly string THUMBPRINT = Environment.GetEnvironmentVariable("THUMBPRINT", EnvironmentVariableTarget.Process);
        static readonly string CLIENT_ID = Environment.GetEnvironmentVariable("CLIENT_ID", EnvironmentVariableTarget.Process);
        static readonly string TENANT = Environment.GetEnvironmentVariable("TENANT", EnvironmentVariableTarget.Process); 

        [FunctionName("ExtractAndApply")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("PnPProvisioningFunction-ExtractAndApply started");


            // Get request body
            string jsonContent = await req.Content.ReadAsStringAsync();
            dynamic data = JsonConvert.DeserializeObject(jsonContent);
            string sourceUrl = data?.sourceUrl;
            string targetUrl = data?.targetUrl;


            //Return 400 if we're missing body params
            if (sourceUrl == null || targetUrl == null)
            {
                return req.CreateResponse(HttpStatusCode.BadRequest, "Please pass a sourceUrl and targetUrl in the request body");
            }

            
            // Get connection certificate details 
            X509Certificate2 cert = null;
            try
            {
                cert = GetCertificate();
            }
            catch (Exception ex)
            {
                log.Error($"CERTIFICATE ERROR: {ex.Message}", ex);
                return req.CreateErrorResponse(HttpStatusCode.InternalServerError, $"CERTIFICATE ERROR: {ex.Message}");
            }

            
            // Get template from source site
            ProvisioningTemplate pnpTemplate = null;
            try
            {
                pnpTemplate = await GetProvisioningTemplate(sourceUrl, cert, log);
            }
            catch (Exception ex)
            {
                log.Error($"EXTRACT ERROR: {ex.Message}", ex);
                return req.CreateErrorResponse(HttpStatusCode.InternalServerError, $"EXTRACT ERROR: {ex.Message}");
            }

            
            // Apply template to target site
            try
            {
                await ApplyProvisioningTemplate(targetUrl, cert, pnpTemplate, log);
            }
            catch (Exception ex)
            {
                log.Error($"APPLY ERROR: {ex.Message}", ex);
                return req.CreateErrorResponse(HttpStatusCode.InternalServerError, $"APPLY ERROR: {ex.Message}");
            }


            // Return 200
            return req.CreateResponse(HttpStatusCode.OK);
        }

        private static X509Certificate2 GetCertificate()
        {
            X509Store certStore = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            certStore.Open(OpenFlags.ReadOnly);
            X509Certificate2Collection certCollection = certStore.Certificates.Find(X509FindType.FindByThumbprint, THUMBPRINT, false);
            return certCollection[0];
        }

        private static async Task<ProvisioningTemplate> GetProvisioningTemplate(string url, X509Certificate2 cert, TraceWriter log)
        {
            ProvisioningTemplate template = null;
            OfficeDevPnP.Core.AuthenticationManager authMgr = new OfficeDevPnP.Core.AuthenticationManager();
            using (ClientContext sourceCtx = authMgr.GetAzureADAppOnlyAuthenticatedContext(url, CLIENT_ID, TENANT, cert))
            {
                // Disable request timeout for extraction
                sourceCtx.RequestTimeout = Timeout.Infinite;

                // Make sure our context is valid
                sourceCtx.Load(sourceCtx.Web, w => w.Url, w => w.Title);
                await sourceCtx.ExecuteQueryRetryAsync();
                log.Info($"Connected to sourceUrl: {sourceCtx.Web.Url}");

                // Extract template and hold in-memory
                ProvisioningTemplateCreationInformation ptci = new ProvisioningTemplateCreationInformation(sourceCtx.Web);
                ptci.ProgressDelegate = delegate (string message, Int32 progress, Int32 total)
                {
                    log.Info(string.Format("EXTRACT: {0:00}/{1:00} - {2}", progress, total, message));
                };

                log.Info($"Beginning PnP template extraction");
                template = sourceCtx.Web.GetProvisioningTemplate(ptci);
                log.Info($"Finished PnP template extraction");
            }
            return await Task.FromResult<ProvisioningTemplate>(template);
        }

        private static async Task<Int32> ApplyProvisioningTemplate(string url, X509Certificate2 cert, ProvisioningTemplate template, TraceWriter log)
        {
            OfficeDevPnP.Core.AuthenticationManager authMgr = new OfficeDevPnP.Core.AuthenticationManager();
            using (ClientContext targetCtx = authMgr.GetAzureADAppOnlyAuthenticatedContext(url, CLIENT_ID, TENANT, cert))
            {
                // Disable request timeout for extraction
                targetCtx.RequestTimeout = Timeout.Infinite;

                // Make sure our context is valid
                targetCtx.Load(targetCtx.Web, w => w.Url, w => w.Title);
                await targetCtx.ExecuteQueryRetryAsync();
                log.Info($"Connected to targetUrl: {targetCtx.Web.Url}");

                // Apply template, clear existing nav nodes
                ProvisioningTemplateApplyingInformation ptai = new ProvisioningTemplateApplyingInformation();
                ptai.ClearNavigation = true;
                ptai.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                {
                    log.Info(string.Format("APPLY: {0:00}/{1:00} - {2}", progress, total, message));
                };

                log.Info($"Beginning applying PnP template");
                targetCtx.Web.ApplyProvisioningTemplate(template, ptai);
                log.Info($"Finished applying PnP template");
            }
            return await Task.FromResult(0);
        }
    }
}
