# Dynamics365Reports
Current WIP, A report framework for Dynamics 365 using an Azure Function to use a HTML layout and FetchXML to create a PDF report.

Currently Planned Features:
* Better credential handeling
* Full sample HTML resources (Templates)
* Dynamics workflow plugin
* Dynamics html resource for Email Report and Print Report buttons
* Extra formatting options

## Usage
1. Create a new HTTP Trigger C# function in Azure. Create / Replace function.json, project.json and run.cx with files under Dynamics365PDF-Report folder.
2. Create a HTML resource in Dynamics 365, use comment synax below to as placeholders for your data.
3. Call the function with the following body syntax (You can also use a query string with the same params):

**Note** The FetchXML is HTML encoded. [You can use website like to HTML encode](https://www.url-encode-decode.com/)

```JSON
{
    "apiKey":"updateTheKeyInRun.csAndPasteHere",
    "crmOrgURI":"https://yourinstance.api.yourcrmregion.dynamics.com/XRMServices/2011/Organization.svc",
    "crmUser":"your@username.here",
    "crmPass":"somePassw0rd!",
    "HTMLResource":"nameOfHTMLResourceInCRM",
    "FetchXML":"&#x3C;fetch version=&#x22;1.0&#x22; output-format=&#x22;xml-platform&#x22; mapping=&#x22;logical&#x22; distinct=&#x22;false&#x22;&#x3E; &#x3C;entity name=&#x22;contact&#x22;&#x3E;    &#x3C;attribute name=&#x22;fullname&#x22; /&#x3E;    &#x3C;attribute name=&#x22;contactid&#x22; /&#x3E;    &#x3C;order attribute=&#x22;fullname&#x22; descending=&#x22;false&#x22; /&#x3E;    &#x3C;filter type=&#x22;and&#x22;&#x3E;      &#x3C;condition attribute=&#x22;contactid&#x22; operator=&#x22;eq&#x22; uiname=&#x22;365&#xA0;Test&#xA0;2&#x22; uitype=&#x22;contact&#x22; value=&#x22;{7FB57B71-8405-E811-8154-E0071B670E51}&#x22; /&#x3E;    &#x3C;/filter&#x3E;  &#x3C;/entity&#x3E;&#x3C;/fetch&#x3E;"
}
```

#### Parameters

*apiKey (string)* - The API key as specified in the C# Azure Function.

*crmOrgURI (string)* - The REST discovery URL for your instance.

*crmUser (string)* - CRM credential username.

*crmPass (string)* - CRM credential password.

*HTMLResource (string)* - The name of the HTML resource to use as the template from CRM.

*FetchXML (string)* - HTML Encoded FetchXML for the subreport data. If you set FetchXML to "Embedded" it will look in the HTML for the FetchXML. See Embedded FetchXML for more info.

## HTML Comment Syntax
1. Start all placeholders with <!--XRMREPORT:
3. Paste in JSON with the options for the placeholder type you would like to use, see Placeholder Types for more info.
2. End all placeholders with -->

### Example
```HTML
<html>
<body>
Name: <!--XRMREPORT:{"Type": "Field", "Field": "new_name"}-->
</body>
</html>
```

## Placeholder Types
### Single Field
Used to insert a single field from the main FetchXML query.
#### Parameters

*Type (string)* - The type of placeholder, in this case "Field"

*Field (string)* - The name of the field to lookup.

**Example**
```HTML
<html>
<body>
Name: <!--XRMREPORT:{"Type": "Field", "Field": "new_name"}-->
</body>
</html>
```
### Checkbox Field
Used to insert a single field from the main FetchXML query as a checkbox. Checkbox will be ticked if the field value matchs the specified CheckedValue.
#### Parameters

*Type (string)* - The type of placeholder, in this case "Checkbox"

*Field (string)* - The name of the field to lookup.

*CheckedValue (string)* - The checkbox will be checked when this the field value matches this parameter.

**Example**
```HTML
<html>
<body>
Name: <!--XRMREPORT:{"Type": "Checkbox", "Field": "new_YesNoField", "CheckedValue" : "Yes"}-->
</body>
</html>
```

### Currency Field
Used to insert a single field from the main FetchXML query and specify the Currency format.
#### Parameters

*Type (string)* - The type of placeholder, in this case "Currency"

*Field (string)* - The name of the field to lookup.

*Culture (string)* - The culture to use for the currency format, eg en-US

**Example**
```HTML
<html>
<body>
Amount: <!--XRMREPORT:{"Type": "Currency", "Field": "new_amount", "Culture" : "en-AU"}-->
</body>
</html>
```
### DateTime Field
Used to insert a single field from the main FetchXML query and specify the Currency format.
#### Parameters

*Type (string)* - The type of placeholder, in this case "DateTime"

*Field (string)* - The name of the field to lookup.

*Culture (string)* - The culture to use for the DateTime format, eg en-US [Full List Here](https://msdn.microsoft.com/en-us/library/cc233982.aspx)

*Format (string)* - The format to use for the DateTime, eg dd/mm/yyyy [Full List Here](https://docs.microsoft.com/en-us/dotnet/standard/base-types/custom-date-and-time-format-strings)

**Example**
```HTML
<html>
<body>
Name: <!--XRMREPORT:{"Type": "DateTime", "Field": "new_amount", "Culture" : "en-AU", "Format" : "dd/mm/yyyy", "OffsetHours" : "0"}-->
</body>
</html>
```

### Subreport Field
Used to insert a Subreport. NB: At this stage the FetchXML is static, planning to allow for XML to be customised on the fly.
#### Parameters

*Type (string)* - The type of placeholder, in this case "Subreport"

*FetchXML (string)* - HTML Encoded FetchXML for the subreport data. If you set FetchXML to "Embedded" it will look in the HTML for the FetchXML.

*Webresource (string)* - Name of the subreport's webresource in CRM.

**Example**
```HTML
<html>
<body>
<table>
<tr><th>Field1</th><th>Field2</th></tr>
<!--XRMREPORT:{"Type": "Subreport", "FetchXML": "htmlEncodedFetchXML", "Webresource" : "new_webresource"}-->
</table>
</body>
</html>
``` 
## Embedded FetchXML
Rather than passing the FetchXML with the request you can embedded it into the HTML report. To do this place a HTML comment at the top of the file containing the XML like follows:

**Example**
```HTML
<!--XRMXML
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
  <entity name="Contact">
    <attribute name="fullname" />
    <filter type="and">
      <condition attribute="contactid" operator="eq" uiname="Test" uitype="contact" value="{03233A15-33FF-E711-8143-70106FA11B81}" />
    </filter>
  </entity>
</fetch>
XRMXML-->
<html>
<body>
Some HTML stuff.
</body>
</html>
``` 