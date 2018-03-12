using Newtonsoft.Json;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System.Text.RegularExpressions;

public static string processHTML (string HTMLResource, OrganizationServiceProxy OSP, string FetchXML){
        string htmlResult = "";
        string getEntValue(string CRMField, Entity ent)
        {
            return ent.FormattedValues.ContainsKey(CRMField) ? ent.FormattedValues[CRMField].ToString() : ent.Attributes.ContainsKey(CRMField) ? ent.Attributes[CRMField].ToString() : "";
        }
        var fetchResult = GetRelatedEntityFromFetchXML(FetchXML, OSP);
        foreach (var EntityRecord in fetchResult.Entities)
        {
            string htmlToProcess = retrieveWebResource(HTMLResource, OSP);
            string pattern = @"<!--XRMREPORT:[^>]*-->";
            Match htmlReportComments = Regex.Match(htmlToProcess, pattern);
            while (htmlReportComments.Success) {
                string jsonData = htmlReportComments.Value.Replace("<!--XRMREPORT:","").Replace("-->","");
                Dictionary<string, string> item = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonData);
                // Replace options
                // Single field from source record <!--XRMREPORT:{"Type": "Field", "Field": "new_name"}-->
                // Checkbox field from source record <!--XRMREPORT:{"Type": "Checkbox", "Field": "new_name", "CheckedValue" : "Yes"}-->
                // Subreport <!--XRMREPORT:{"Type": "Subreport", "FetchXML": "htmlEncodedFetchXML", "Webresource" : "new_webresource"}-->
                switch (item["Type"].ToString())
                {
                    case "Field":
                        htmlToProcess = htmlToProcess.Replace(htmlReportComments.Value, getEntValue(item["Field"],EntityRecord));
                        break;
                    case "Checkbox":
                        htmlToProcess = htmlToProcess.Replace(htmlReportComments.Value, getEntValue(item["Field"], EntityRecord) == item["CheckedValue"] ? "<input type='checkbox' checked>" : "<input type='checkbox'>" );
                        break;
                    case "Subreport":
                        string SubreportFetchXML = item["FetchXML"];
                        htmlToProcess = htmlToProcess.Replace(htmlReportComments.Value, processHTML(item["Webresource"], OSP, SubreportFetchXML));
                        break;
                    default:
                        htmlToProcess = htmlToProcess.Replace(htmlReportComments.Value,"<!--XRMREPORT: No Match-->");
                        break;
                }
                htmlReportComments = htmlReportComments.NextMatch();
            }
            htmlResult += htmlToProcess;   
        }
        return htmlResult;
    }

//Get related entity via fetchXML
public static  EntityCollection GetRelatedEntityFromFetchXML(string fetchXML, OrganizationServiceProxy OSP){
        EntityCollection tempDCEntity = OSP.RetrieveMultiple(new FetchExpression(fetchXML));
         return tempDCEntity;
    }

// import HTML templates
public static string retrieveWebResource(string webresourceName, OrganizationServiceProxy OSP) {
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
        EntityCollection webResourceEntityCollection = OSP.RetrieveMultiple(requestWebResource);

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