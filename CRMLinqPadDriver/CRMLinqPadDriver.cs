/*================================================================================================================================

  This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.  

  THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
  INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  

  We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object 
  code form of the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks to market Your software 
  product in which the Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in which the 
  Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims 
  or lawsuits, including attorneys’ fees, that arise or result from the use or distribution of the Sample Code.

 =================================================================================================================================*/

using LINQPad.Extensibility.DataContext;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Pfe.Xrm.Helper;
using Microsoft.Pfe.Xrm.View;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Microsoft.Pfe.Xrm
{
    /// <summary>
    /// Create StaticDataContextDriver, though it is dynamically generated. By using Static, it is easy to reuse generated assembly.
    /// </summary>
    public class CRMLinqPadDriver : StaticDataContextDriver
    {
        public OrganizationServiceEx orgService;

        // Author
        public override string Author
        {
            get { return "Dynamics CRM PFE"; }
        }

        // Display Name for connection.
        public override string GetConnectionDescription(IConnectionInfo cxInfo)
        {
            CrmProperties props = new CrmProperties(cxInfo);
            return props.FriendlyName + " " + props.OrgUri;
        }

        // Display Name for Driver.
        public override string Name
        {
            get { return "Dynamics CRM Linq Pad Driver"; }
        }

        /// <summary>
        /// Opens Login dialog
        /// </summary>
        /// <param name="cxInfo">IConnectionInfo</param>
        /// <param name="isNewConnection">Indicate if this is new connection request or update existing connection</param>
        /// <returns>result</returns>
        public override bool ShowConnectionDialog(IConnectionInfo cxInfo, bool isNewConnection)
        {
            return new MainWindow(cxInfo, isNewConnection).ShowDialog() == true;
        }

        /// <summary>
        /// Additional work for Initialization. Hook PreExecute event to display QueryExpression and FetchXML to SQL tab.
        /// </summary>
        /// <param name="cxInfo">IConnectionInfo</param>
        /// <param name="context">XrmContext</param>
        /// <param name="executionManager">QueryExecutionManager</param>
        public override void InitializeContext(IConnectionInfo cxInfo, object context, QueryExecutionManager executionManager)
        {
            // Attach PreExecute event.
            orgService.PreExecute += (s, e) => 
            {
                // Take QueryExpression and convert it to FetchXML.
                QueryExpressionToFetchXmlRequest request = new QueryExpressionToFetchXmlRequest { Query = e.query };
                string fetchXml = (orgService.Execute(request) as QueryExpressionToFetchXmlResponse).FetchXml;

                // Instantiate DataContractSerializer to serialiez QueryExpression.
                DataContractSerializer serializer = new DataContractSerializer(typeof(QueryExpression));
                using (MemoryStream stream = new MemoryStream())
                {
                    // Use XmlWriter with Indent.
                    using (XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings{ Indent = true}))
                    {
                        serializer.WriteObject(writer, e.query);
                    }

                    // Display result.
                    executionManager.SqlTranslationWriter.WriteLine("QueryExpression:");
                    executionManager.SqlTranslationWriter.WriteLine(Encoding.UTF8.GetString(stream.GetBuffer()));
                    executionManager.SqlTranslationWriter.WriteLine();
                }        
       
                // Display FetchXml as well.
                using (MemoryStream stream = new MemoryStream())
                {
                    XDocument xdoc = XDocument.Parse(fetchXml);
                    using (XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings { Indent = true }))
                    {
                        xdoc.WriteTo(writer);
                    }

                    executionManager.SqlTranslationWriter.WriteLine("FetchXml:");
                    executionManager.SqlTranslationWriter.WriteLine(Encoding.UTF8.GetString(stream.GetBuffer()).Replace("\"", "\'"));
                }
            };            

            base.InitializeContext(cxInfo, context, executionManager);
        }
        
        /// <summary>
        /// Pass Extended OrganizationService to context
        /// </summary>
        /// <param name="cxInfo">IConnectionInfo</param>
        /// <returns>Constructor argument(s)</returns>
        public override object[] GetContextConstructorArguments(IConnectionInfo cxInfo)
        {
            // Instantiate CrmProperties.
            CrmProperties props = new CrmProperties(cxInfo);
            // Create CrmConnection depending on connection type.
            CrmConnection crmConn =null;
            switch (props.AuthenticationProviderType)
            {
                case "OnlineFederation":
                    crmConn = CrmConnection.Parse(String.Format("Url={0}; Username={1}; Password={2};", props.OrgUri, props.UserName, props.Password));
                    break;
                case "ActiveDirectory":
                    if(String.IsNullOrEmpty(props.DomainName))
                        crmConn = CrmConnection.Parse(String.Format("Url={0};", props.OrgUri));
                    else
                        crmConn = CrmConnection.Parse(String.Format("Url={0}; Domain={1}; Username={2}; Password={3};", props.OrgUri, props.DomainName, props.UserName, props.Password));
                    break;
                case "Federation":
                    crmConn = CrmConnection.Parse(String.Format("Url={0}; Username={1}; Password={2};", props.OrgUri, props.UserName, props.Password));
                    break;
            }
            // Instantiate Extended OrganizationService.
            orgService = new OrganizationServiceEx(crmConn);
            
            return new object[]
            {
                new OrganizationService(orgService)
            };             
        }

        /// <summary>
        /// Specify Context Constructor argument type(s)
        /// </summary>
        /// <param name="cxInfo">IConnectionInfo</param>
        /// <returns>Constructor argument type(s)</returns>
        public override ParameterDescriptor[] GetContextConstructorParameters(IConnectionInfo cxInfo)
        {
            return new ParameterDescriptor[]
            {
                // OrgainzationService is the only constructor argument.
                new ParameterDescriptor("OrganizationService", "Microsoft.Xrm.Client.Services.OrganizationService")
            };
        }

        /// <summary>
        /// Load additional assemblies to LinqPad process.
        /// </summary>
        /// <returns>List of assmeblies to be loaded.</returns>
        public override IEnumerable<string> GetAssembliesToAdd(IConnectionInfo cxInfo)
        {
            return new string[]
            {
                "System.Data.Services.dll",
                "System.ServiceModel.dll",
                "System.Runtime.Serialization.dll",
                "Microsoft.Crm.Sdk.Proxy.dll",
                "Microsoft.Xrm.Sdk.dll",
                "Microsoft.Xrm.Client.dll"
            };
        }

        /// <summary>
        /// Import additional namespaces.
        /// </summary>
        /// <returns></returns>
        public override IEnumerable<string> GetNamespacesToAdd(IConnectionInfo cxInfo)
        {
            return new string[]
            {
                "System.ServiceModel",
                "System.ServiceModel.Description",
                "Microsoft.Xrm.Sdk",
                "Microsoft.Xrm.Sdk.Query",
                "Microsoft.Xrm.Sdk.Client",
                "Microsoft.Xrm.Sdk.Messages",
                "Microsoft.Xrm.Sdk.Metadata",
                "Microsoft.Crm.Sdk.Messages",
                "Microsoft.Xrm.Sdk.Discovery",
                "Microsoft.Xrm.Client",
                "Microsoft.Xrm.Client.Services"
            };
        }

        /// <summary>
        /// Generate Schema information.
        /// </summary>
        /// <param name="cxInfo">IConnectionInfo</param>
        /// <param name="customType">Context Type</param>
        /// <returns>Schema Information.</returns>
        public override List<ExplorerItem> GetSchema(IConnectionInfo cxInfo, Type customType)
        {
            // Instantiate CrmProperties.
            CrmProperties props = new CrmProperties(cxInfo);
            // Instantiate ExplorerItem list.
            List<ExplorerItem> schema = new List<ExplorerItem>();

            // Load assembly for this connection, as context doesn't have all necessary types.
            Assembly assembly = Assembly.LoadFile(props._cxInfo.CustomTypeInfo.CustomAssemblyPath);

            // Find context first. (You can use customType, too)
            var context = assembly.DefinedTypes.Where(x => x.Name == "XrmContext").First();

            // Create Tables for Schema.
            foreach (PropertyInfo prop in context.DeclaredProperties)
            {
                // If property is not Generic Type or IQueryable, then ignore.
                if (!prop.PropertyType.IsGenericType || prop.PropertyType.Name != "IQueryable`1")
                    continue;

                // Create ExploreItem with Table icon as this is top level item.
                ExplorerItem item = new ExplorerItem(prop.Name, ExplorerItemKind.QueryableObject, ExplorerIcon.Table)
                {
                    // Store Entity Type to Tag.
                    Tag = prop.PropertyType.GenericTypeArguments[0].Name,
                    IsEnumerable = true,
                    Children = new List<ExplorerItem>()
                };

                schema.Add(item);
            }

            // Then create columns for each table. Loop through tables again.
            foreach (PropertyInfo prop in context.DeclaredProperties)
            {
                // Obtain Table item from table lists.
                var item = schema.Where(x => x.Text == prop.Name).First();

                // Get all property from Entity for the table. (i.e. Account for AccountSet)
                foreach (PropertyInfo childprop in assembly.DefinedTypes.Where(x => x.Name == prop.PropertyType.GenericTypeArguments[0].Name).First().DeclaredProperties)
                {
                    // If property is IEnumerable type, then it is 1:N or N:N relationship field.
                    // Need to find a way to figure out if this is 1:N or N:N. At the moment, I just make them as OneToMany type.
                    if (childprop.PropertyType.IsGenericType && childprop.PropertyType.Name == "IEnumerable`1")
                    {
                        // Try to get LinkTarget. 
                        ExplorerItem linkTarget = schema.Where(x => x.Tag.ToString() == childprop.PropertyType.GetGenericArguments()[0].Name).FirstOrDefault();
                        if (linkTarget == null)
                            continue;

                        // Create ExplorerItem as Colleciton Link.
                        item.Children.Add(
                            new ExplorerItem(
                                childprop.Name,
                                ExplorerItemKind.CollectionLink,
                                ExplorerIcon.OneToMany)
                            {
                                HyperlinkTarget = linkTarget,
                                ToolTipText = DataContextDriver.FormatTypeName(childprop.PropertyType, false)
                            });
                    }
                    else
                    {
                        // Try to get LinkTarget to check if this field is N:1.
                        ExplorerItem linkTarget = schema.Where(x => x.Tag.ToString() == childprop.PropertyType.Name).FirstOrDefault();

                        // If no linkTarget exists, then this is normal field.
                        if (linkTarget == null)
                        {
                            // Create ExplorerItem as Column.
                            item.Children.Add(
                                new ExplorerItem(
                                    childprop.Name + " (" + DataContextDriver.FormatTypeName(childprop.PropertyType, false) + ")",
                                    ExplorerItemKind.Property,
                                    ExplorerIcon.Column)
                                {
                                    ToolTipText = DataContextDriver.FormatTypeName(childprop.PropertyType, false)
                                });
                        }
                        else
                        {
                            // Otherwise, create ExploreItem as N:1
                            item.Children.Add(
                                new ExplorerItem(
                                    childprop.Name + " (" + DataContextDriver.FormatTypeName(childprop.PropertyType, false) + ")",
                                    ExplorerItemKind.ReferenceLink,
                                    ExplorerIcon.ManyToOne)
                                {
                                    HyperlinkTarget = linkTarget,
                                    ToolTipText = DataContextDriver.FormatTypeName(childprop.PropertyType, false)
                                });
                        }
                    }
                }
            }

            return schema;
        }
    }
}
