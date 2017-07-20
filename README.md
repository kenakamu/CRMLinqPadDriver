# Use LINQPad to run LINQ query against Microsoft Dynamics CRM
This project is hosted in CodePlex but as it is closing down, I moved to here. 

## Project Description
This is LINQPad driver which enables you to connect to your Microsoft Dynamics CRM organization and run LINQ queries quickly. In addition, you are able to write and execute C# code by using this drive as it supports OrganizationService common methods as well.

## How to use this
Please find detail instruction how to use this driver at LINQPad 4 Driver for Dynamics CRM is available on CodePlex@CRM in the Field Blog

## Feedback
Please let us know what's working, what isn't, as well as suggestions for new capabilities by submitting a new issue

## Connect
In addition to providing feedback on this project site, we'd love to hear directly from you! See what we're up to via our MSDN blog, CRM in the Field or follow us on Twitter: @pfedynamics

## Additional Information
This solution uses CrmOrganizationServiceContext for LINQ capability, which uses SOAP and displays corresponding QueryExpression and FetchXML. As this endpoint is deprecated with latest version of Dynamics 365, please use the Web API version instead. (CRMLinqPadDriverWebAPI)[https://github.com/kenakamu/CRMLinqPadDriverWebAPI]
