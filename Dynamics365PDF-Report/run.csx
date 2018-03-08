#r "Newtonsoft.Json"
//#load "file.csx"

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net;
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
    string EntityID = req.GetQueryNameValuePairs().FirstOrDefault(q => string.Compare(q.Key, "EntityID", true) == 0).Value;
    string EntityName = req.GetQueryNameValuePairs().FirstOrDefault(q => string.Compare(q.Key, "EntityName", true) == 0).Value;
    string FetchXML = req.GetQueryNameValuePairs().FirstOrDefault(q => string.Compare(q.Key, "FetchXML", true) == 0).Value;
    string HTMLResource = req.GetQueryNameValuePairs().FirstOrDefault(q => string.Compare(q.Key, "HTMLResource", true) == 0).Value;
    string crmOrgURI = req.GetQueryNameValuePairs().FirstOrDefault(q => string.Compare(q.Key, "crmOrgURI", true) == 0).Value;
    string apiKey = req.GetQueryNameValuePairs().FirstOrDefault(q => string.Compare(q.Key, "apiKey", true) == 0).Value;
    bool Attachment = true;
    if (crmOrgURI == null){
        log.Info("Using Request Body");
        string bodyContent = await req.Content.ReadAsStringAsync();
        Dictionary<string, string> htmlAttributes = JsonConvert.DeserializeObject<Dictionary<string, string>>(bodyContent);
        EntityID = htmlAttributes["EntityID"];
        EntityName = htmlAttributes["EntityName"];
        FetchXML = htmlAttributes["FetchXML"];
        HTMLResource = htmlAttributes["HTMLResource"];
        crmOrgURI = htmlAttributes["crmOrgURI"];
        apiKey = htmlAttributes["apiKey"];
        Attachment = false;
    } else
    {
        log.Info("Using Query");
    }
    //check if key is correct
    if (apiKey != "dfgfdgfdgfdgavasdsadvx979amdb2349sbkmb435skvd430") { 
        return new HttpResponseMessage(HttpStatusCode.Forbidden);
    } 
    
    // connect to CRM
    IServiceManagement<IOrganizationService> orgServiceManagement = ServiceConfigurationFactory.CreateManagement<IOrganizationService>(new Uri("https://"+crmOrg+".api.crm6.dynamics.com/XRMServices/2011/Organization.svc"));    
    AuthenticationCredentials authCredentials = new AuthenticationCredentials();
    authCredentials.ClientCredentials.UserName.UserName = "";
    authCredentials.ClientCredentials.UserName.Password = "";
    AuthenticationCredentials tokenCredentials = orgServiceManagement.Authenticate(authCredentials);
    OrganizationServiceProxy organizationProxy = new OrganizationServiceProxy(orgServiceManagement, tokenCredentials.SecurityTokenResponse);

    // import HTML templates
    // need to add support for looking up HTML webresources
    string retrieveWebResource(string webresourceName)
        {
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
                string resourceContent = Encoding.UTF8.GetString(binary);
                string byteOrderMarkUtf8 = Encoding.UTF8.GetString(Encoding.UTF8.GetPreamble());                
                if (resourceContent.StartsWith("\""))
                {
                    resourceContent = resourceContent.Remove(0, byteOrderMarkUtf8.Length);
                }               
            }
            return resourceContent;
        }

    string htmlTemplate = retrieveWebResource(HTMLResource);
   //functions

    //Get related entity via Query Expression
    DataCollection<Entity> GetRelatedEntityFromQuery(string retriveEntity, string filterAttribute, Guid filterValue,string orderField, OrderType orderType)
    {
        QueryExpression tempQuery = new QueryExpression
        {
            EntityName = retriveEntity, ColumnSet = new ColumnSet(true), 
            Criteria = new FilterExpression {
                Conditions = {
                    new ConditionExpression {
                        AttributeName = filterAttribute,
                        Operator = ConditionOperator.Equal,
                        Values = { filterValue }
                    }
                }
            },                
            Orders = {
                    new OrderExpression {
                        AttributeName = orderField,
                        OrderType = orderType
                    }
            }
        };
        DataCollection<Entity> tempDCEntity = organizationProxy.RetrieveMultiple(tempQuery).Entities;
        return tempDCEntity;
    }
    //Get related entity via Query Expression   
    EntityCollection GetRelatedEntityFromFetchXML(string fetchXML){
         return organizationProxy.RetrieveMultiple(new FetchExpression(fetchXML));
    }

    //Get vaule of field from entity
    string getEntValue(string CRMField, Entity ent)
    {
        return ent.FormattedValues.ContainsKey(CRMField) ? ent.FormattedValues[CRMField].ToString() : ent.Attributes.ContainsKey(CRMField) ? ent.Attributes[CRMField].ToString() : "";
    }

    // check method for base entity
    // Get CRM data for current entity
    var EntityRecord = new DataCollection<Entity>;
    if (EntityID != null && EntityName != null && FetchXML == null && crmOrgURI != null  && HTMLResource != null){
            log.Info("Using EntityID");
            var CRMColumnSetAll = new ColumnSet(true);
            var EntityGuid = new Guid(EntityID);
            EntityRecord = organizationProxy.Retrieve(EntityName,EntityGuid,CRMColumnSetAll);
    } else if (EntityID == null && EntityName == null && FetchXML != null && crmOrgURI != null  && HTMLResource != null){
            log.Info("Using FetchXML");
            EntityRecord = GetRelatedEntityFromFetchXML(FetchXML);
    } else { 
        return new HttpResponseMessage(HttpStatusCode.BadRequest);
    } 
 
    //replace options
    // Single field from source record <!--XRMREPORT:{"Type": "Field", "Field": "new_name"}-->
    // Single field from related record <!--XRMREPORT:{"Type": "Lookup", "Lookup": "new_contact", "Field": "firstname"}-->
    // WIP Subreport <!--XRMREPORT:{"Type": "Subgrid", "RelatedEntity": "contacts", "RelatedField" : "new_field"}-->

    //replace HTML template fields
    string pattern = @"/<!--XRMREPORT:[^>]*-->/g";
    Match htmlReportComments = Regex.Match(htmlTemplate, pattern);
    while (htmlReportComments.Success) {
        log.Info("Working with comment: " + htmlReportComments.Value);
        dynamic item = JsonConvert.DeserializeObject<dynamic>(htmlReportComments.Value.Replace("<!--XRMREPORT:","").Replace("-->",""));
        switch (item.Type)
        {
            case "Field":
                log.Info("Case Field");
                htmlTemplate = htmlTemplate.Replace(htmlReportComments.Value, getEntValue(item.Field,EntityRecord));
                break;
            case "Lookup":
                log.Info("Case Lookup");
                break;
            default:
                log.Info("No Match");
                break;
        }
        htmlTemplate.Replace(htmlReportComments.Value,htmlReportComments.Value);
        htmlReportComments = htmlReportComments.NextMatch();
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


