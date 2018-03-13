#r "Newtonsoft.Json"
#load "functions.csx"
using Newtonsoft.Json;
using System.Net;
using System.Text;
using System.Web;
using System.Net.Http.Headers;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using OpenHtmlToPdf; 


public static async Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log, ExecutionContext context)
{
    log.Info("C# HTTP trigger function processed a request.");

    // parse query parameter or request body
    string FetchXML = HttpUtility.HtmlDecode(req.GetQueryNameValuePairs().FirstOrDefault(q => string.Compare(q.Key, "FetchXML", true) == 0).Value);
    string HTMLResource = req.GetQueryNameValuePairs().FirstOrDefault(q => string.Compare(q.Key, "HTMLResource", true) == 0).Value;
    string crmOrgURI = req.GetQueryNameValuePairs().FirstOrDefault(q => string.Compare(q.Key, "crmOrgURI", true) == 0).Value;
    string apiKey = req.GetQueryNameValuePairs().FirstOrDefault(q => string.Compare(q.Key, "apiKey", true) == 0).Value;
    string crmUser = req.GetQueryNameValuePairs().FirstOrDefault(q => string.Compare(q.Key, "crmUser", true) == 0).Value;
    string crmPass = req.GetQueryNameValuePairs().FirstOrDefault(q => string.Compare(q.Key, "crmPass", true) == 0).Value;
    string returnType = req.GetQueryNameValuePairs().FirstOrDefault(q => string.Compare(q.Key, "returnType", true) == 0).Value;
    bool Attachment = true;
    if (crmOrgURI == null){
        log.Info("Using Request Body");
        string bodyContent = await req.Content.ReadAsStringAsync();
        Dictionary<string, string> htmlAttributes = JsonConvert.DeserializeObject<Dictionary<string, string>>(bodyContent);
        FetchXML = HttpUtility.HtmlDecode(htmlAttributes["FetchXML"]);
        HTMLResource = htmlAttributes["HTMLResource"];
        crmOrgURI = htmlAttributes["crmOrgURI"];
        apiKey = htmlAttributes["apiKey"];
        crmUser = htmlAttributes["crmUser"];
        crmPass = htmlAttributes["crmPass"];
        returnType = htmlAttributes["returnType"];
        Attachment = false;
    } else
    {
        log.Info("Using Query");
    }
    // check if key is correct
    if (apiKey != "dfgfdgfdgfdgavasdsadvx979amdb2349sbkmb435skvd430") { 
        return new HttpResponseMessage(HttpStatusCode.Forbidden);
    } 
    // check parameters
    if (FetchXML == null || crmOrgURI == null  || HTMLResource == null){
        return new HttpResponseMessage(HttpStatusCode.BadRequest);
    } 
    
    // connect to CRM
    IServiceManagement<IOrganizationService> orgServiceManagement = ServiceConfigurationFactory.CreateManagement<IOrganizationService>(new Uri(crmOrgURI));    
    AuthenticationCredentials authCredentials = new AuthenticationCredentials();
    authCredentials.ClientCredentials.UserName.UserName = crmUser;
    authCredentials.ClientCredentials.UserName.Password = crmPass;
    AuthenticationCredentials tokenCredentials = orgServiceManagement.Authenticate(authCredentials);
    OrganizationServiceProxy organizationProxy = new OrganizationServiceProxy(orgServiceManagement, tokenCredentials.SecurityTokenResponse);

    // Process HTML Template
    string htmlTemplate = processHTML(HTMLResource, organizationProxy, FetchXML);
    log.Info(htmlTemplate);
    // Return HTML if returnType = HTML
    if (returnType == "html"){
        return new HttpResponseMessage(HttpStatusCode.OK) {
            Content = new StringContent(htmlTemplate, Encoding.UTF8, "application/text")
        };
    }

    // PDF Creation
    log.Info("Processing PDF Request");
    var pdf = Pdf.From(htmlTemplate).Content();
    var res = new HttpResponseMessage(HttpStatusCode.OK);
        res.Content = new ByteArrayContent(pdf);
        res.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
        if (Attachment)
        {
            res.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
        }else
        {
            res.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("inline");
        }
    return res;

}