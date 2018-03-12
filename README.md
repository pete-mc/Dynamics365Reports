# Dynamics365Reports-AzureFunction
Current WIP, planning to create a report framework for Dynamics 365 using a HTML layout and Azure Function to create a PDF report.

Currently Planned Features:
* Better support for currency formats
* Better support for date/time formats
* Full sample HTML resources (Templates)
* Dynamics workflow plugin
* Dynamics html resource for Email Report and Print Report buttons

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

**Type (string)**

The type of placeholder, in this case "Field"


**Field (string)**

The name of the field to lookup.

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

**Type (string)**

The type of placeholder, in this case "Checkbox"


**Field (string)**

The name of the field to lookup.

**CheckedValue (string)**

The checkbox will be checked when this the field value matches this parameter.

**Example**
```HTML
<html>
<body>
Name: <!--XRMREPORT:{"Type": "Checkbox", "Field": "new_YesNoField", "CheckedValue" : "Yes"}-->
</body>
</html>
```
