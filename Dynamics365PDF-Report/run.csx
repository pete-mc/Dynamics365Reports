#r "Newtonsoft.Json"
//#load "file.csx"

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Net;
using System.Text;
using System.Web;
using System.Net.Http.Headers;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System.Linq;
using System.IO;
using OpenHtmlToPdf; 
using System.Text.RegularExpressions;

public static async Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log, ExecutionContext context)
{
    log.Info("C# HTTP trigger function processed a request.");

    // parse query parameter or request body
    // need to add support for [EntityName, EntityID], [FetchXML] & HTML resources
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
    //check if key is correct
    if (apiKey != "dfgfdgfdgfdgavasdsadvx979amdb2349sbkmb435skvd430") { 
        return new HttpResponseMessage(HttpStatusCode.Forbidden);
    } 
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

    // import HTML templates
    // need to add support for looking up HTML webresources
    string retrieveWebResource(string webresourceName)
        {
            string resourceContent = "";
            ColumnSet cols = new ColumnSet();
            cols.AddColumn("content");
            QueryByAttribute requestWebResource = new QueryByAttribute
            {
                EntityName = "webresource",
                ColumnSet = cols
            };
            requestWebResource.Attributes.AddRange("name");
            requestWebResource.Values.AddRange(webresourceName);

            Entity webResourceEntity = null;
            EntityCollection webResourceEntityCollection = organizationProxy.RetrieveMultiple(requestWebResource);

            if (webResourceEntityCollection.Entities.Count > 0)
            {
                webResourceEntity = webResourceEntityCollection.Entities[0];
                byte[] binary = Convert.FromBase64String(webResourceEntity.Attributes["content"].ToString());
                resourceContent = Encoding.UTF8.GetString(binary);
                string byteOrderMarkUtf8 = Encoding.UTF8.GetString(Encoding.UTF8.GetPreamble());                
                if (resourceContent.StartsWith("\""))
                {
                    resourceContent = resourceContent.Remove(0, byteOrderMarkUtf8.Length);
                }       
                        
            }
            return resourceContent;
        }

    string htmlTemplate = retrieveWebResource(HTMLResource);
    log.Info("Working with html: " + htmlTemplate);

    //Get related entity via Query Expression   
    EntityCollection GetRelatedEntityFromFetchXML(string fetchXML){
        EntityCollection tempDCEntity = organizationProxy.RetrieveMultiple(new FetchExpression(fetchXML));
         return tempDCEntity;
    }

    //Get value of field from entity
    string getEntValue(string CRMField, Entity ent)
    {
        return ent.FormattedValues.ContainsKey(CRMField) ? ent.FormattedValues[CRMField].ToString() : ent.Attributes.ContainsKey(CRMField) ? ent.Attributes[CRMField].ToString() : "";
    }

    // Get CRM data for current entity
    var EntityRecord = GetRelatedEntityFromFetchXML(FetchXML).Entities[0];

    //replace options
    // Single field from source record <!--XRMREPORT:{"Type": "Field", "Field": "new_name"}-->
    // Checkbox field from source record <!--XRMREPORT:{"Type": "Checkbox", "Field": "new_name", "CheckedValue" : "Yes"}-->
    // WIP Subreport <!--XRMREPORT:{"Type": "Subreport", "FetchXML": "<blah>", "Webresource" : "webresource.html"}-->

    //replace HTML template fields
    string pattern = @"<!--XRMREPORT:[^>]*-->";
    Match htmlReportComments = Regex.Match(htmlTemplate, pattern);
    while (htmlReportComments.Success) {
        log.Info("Working with comment: " + htmlReportComments.Value);
        string jsonData = htmlReportComments.Value.Replace("<!--XRMREPORT:","").Replace("-->","");
        log.Info("Json Data: " + jsonData);
        Dictionary<string, string> item = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonData);
        switch (item["Type"].ToString())
        {
            case "Field":
                log.Info("Case Field");
                htmlTemplate = htmlTemplate.Replace(htmlReportComments.Value, getEntValue(item["Field"],EntityRecord));
                break;
            case "Checkbox":
                log.Info("Case Checkbox");
                htmlTemplate = htmlTemplate.Replace(htmlReportComments.Value, getEntValue(item["Field"], EntityRecord) == item["CheckedValue"] ? "<input type='checkbox' checked>" : "<input type='checkbox'>" );
                break;
            case "Webresource":
                log.Info("Case Webresource: Code WIP");
                break;
            default:
                log.Info("No Match");
                htmlTemplate = htmlTemplate.Replace(htmlReportComments.Value,"<!--XRMREPORT: No Match-->");
                break;
        }
        htmlReportComments = htmlReportComments.NextMatch();
      }   
    
    // PDF Creation
    if (returnType == "html"){
        return new HttpResponseMessage(HttpStatusCode.OK) {
            Content = new StringContent(htmlTemplate, Encoding.UTF8, "application/text")
        };
    }

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


