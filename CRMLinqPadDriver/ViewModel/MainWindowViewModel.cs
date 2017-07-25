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
using Microsoft.Pfe.Xrm.Common;
using Microsoft.Pfe.Xrm.View;
using Microsoft.CSharp;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.ServiceModel.Description;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Xrm.Sdk.Discovery;

namespace Microsoft.Pfe.Xrm.ViewModel
{
    public class MainWindowViewModel : ViewModelBase
    {
        #region Property

        // Indicate if Entites are loaded or not.
        private bool isNewConnection;
        public bool IsNewConnection
        {
            get { return isNewConnection; }
            set
            {
                isNewConnection = value;
                NotifyPropertyChanged();
            }
        }

        // Message holder
        private string loadMessage;
        public string LoadMessage
        {
            get { return loadMessage; }
            set
            {
                loadMessage = value;
                NotifyPropertyChanged();
            }
        }

        // Message holder
        private string message;
        public string Message
        {
            get { return message; }
            set
            {
                message = value;
                NotifyPropertyChanged();
            }
        }

        // CrmProperties
        private CrmProperties props;

        // Indicate if schema as been loaded.
        private bool isLoaded;
        public bool IsLoaded
        {
            get { return isLoaded; }
            set
            {
                isLoaded = value;
                NotifyPropertyChanged();
            }
        }

        #endregion

        #region Method

        /// <summary>
        /// Constructor
        /// </summary>
        public MainWindowViewModel(IConnectionInfo cxInfo, bool isNewConnection)
        {
            // Display message depending on isNewConnection
            Message = isNewConnection ? "Click Login button to Login." : "Click Reload Data to update Schema";

            IsLoaded = false;
            props = new CrmProperties(cxInfo);
            IsNewConnection = isNewConnection;
        }        

        /// <summary>
        /// Raised when the login form process is completed.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ctrl_ConnectionToCrmCompleted(object sender, EventArgs e)
        {
            if (sender is CrmLogin)
            {
                ((CrmLogin)sender).Close();
            }
        }

        /// <summary>
        /// Generate Context and Early bound class file by using CrmSvcUtil.exe.
        /// Then compile it into an assembly.
        /// </summary>
        private void LoadData()
        {
            AppDomain.CurrentDomain.AssemblyResolve += (o, args) =>
            {
                var name = new AssemblyName(args.Name);
                return name.Name == "Microsoft.Xrm.Sdk" ? typeof(Microsoft.Xrm.Sdk.Entity).Assembly : null;
            };
            // Generate code by using CrmSvcUtil.exe
            string code = GenerateCode(props);
            
            // Store assembly full path.
            string assemblyFullName = "";
            // When reload data, we need to generate new assembly as current one is hold by LinqPad instance.
            // Switch Context.dll and ContextAlternative.dll based on what LinqPad loads right now.
            if (props._cxInfo.CustomTypeInfo.CustomAssemblyPath.EndsWith("Alternative.dll"))
                assemblyFullName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, props.OrgUri.GetHashCode() + "Context.dll");
            else
                assemblyFullName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, props.OrgUri.GetHashCode() + "ContextAlternative.dll");

            props._cxInfo.CustomTypeInfo.CustomAssemblyPath = assemblyFullName;

            // Compile the code into the assembly. To avoid duplicate name for each connection, hash entire URL to make it unique.
            BuildAssembly(code, assemblyFullName);

            // Then delete generated file.
            File.Delete(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "XrmData.cs"));

            // Update message.
            Message = "Loading Complete. Click Exit and wait a while until Linq Pad displays full Schema information.";
            IsLoaded = true;
        }

        /// <summary>
        /// Generate context for LinqPad by using CrmSvcUtil.exe. In this example, we generate CrmOrganizationServiceContext
        /// so that context has all common methods from OrganizationService. However if you prefer to have just Linq capabilities, 
        /// you are also able to generate OrganizationServiceContext only. Please refer to https://msdn.microsoft.com/en-us/library/gg695792.aspx
        /// for more detail about CrmSvcUtil too.
        /// </summary>
        /// <param name="props">CRM Properties</param>
        /// <returns>Generate code</returns>
        private string GenerateCode(CrmProperties props)
        {            
            // Create Process
            Process p = new Process();

            p.StartInfo.UseShellExecute = false;
            // Specify CrmSvcUtil.exe as process name.
            p.StartInfo.FileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CrmSvcUtil.exe");
            // Do not display window
            p.StartInfo.CreateNoWindow = true;
            p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            // Assign argumemnt for CrmSvcUtil. This format works for all environment.
            p.StartInfo.Arguments =
                String.Format("/codeCustomization:\"Microsoft.Xrm.Client.CodeGeneration.CodeCustomization, Microsoft.Xrm.Client.CodeGeneration\" /url:{0} /username:{1} /password:{2} /out:\"{3}\" /namespace:Microsoft.Pfe.Xrm /serviceContextName:XrmContext",
                props.OrgUriActual,
                props.UserName,
                props.Password,
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "XrmData.cs"));

            // Execute and wait until it complited.
            p.Start();
            p.WaitForExit();

            // Read generate file and return it.
            return System.IO.File.ReadAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "XrmData.cs"));
        }

        /// <summary>
        /// Generate an assembly
        /// </summary>
        /// <param name="code">code to be compiled</param>
        /// <param name="name">assembly name</param>
        private void BuildAssembly(string code, string name)
        {
            // Use the CSharpCodeProvider to compile the generated code:
            CompilerResults results;
            using (var codeProvider = new CSharpCodeProvider(new Dictionary<string, string>() { { "CompilerVersion", "v4.0" } }))
            {
                var options = new CompilerParameters(
                    "System.dll System.Core.dll System.Xml.dll System.Data.Services.Client.dll System.Runtime.Serialization.dll System.Data.Services.dll".Split(' '),
                    name,
                    false);
                // Force load Microsoft.Xrm.Sdk assembly.
                options.ReferencedAssemblies.Add(typeof(Microsoft.Xrm.Sdk.Entity).Assembly.Location);
                options.ReferencedAssemblies.Add(typeof(Microsoft.Xrm.Client.CrmEntity).Assembly.Location);

                // Compile
                results = codeProvider.CompileAssemblyFromSource(options, code);
            }
            if (results.Errors.Count > 0)
                throw new Exception
                    ("Cannot compile typed context: " + results.Errors[0].ErrorText + " (line " + results.Errors[0].Line + ")");
        }

        #endregion

        #region Command

        /// <summary>
        /// Login by using CrmLogin and generate an assembly
        /// </summary>
        public RelayCommand LoginCommand
        {
            get
            {
                return new RelayCommand(async () =>
                {
                    // Establish the Login control
                    CrmLogin ctrl = new CrmLogin();
                    // Wire Event to login response. 
                    ctrl.ConnectionToCrmCompleted += ctrl_ConnectionToCrmCompleted;
                    // Show the dialog. 
                    ctrl.ShowDialog();

                    // Handel return. 
                    if (ctrl.CrmConnectionMgr != null && ctrl.CrmConnectionMgr.CrmSvc != null && ctrl.CrmConnectionMgr.CrmSvc.IsReady)
                    {
                        IsLoading = true;
                        LoadMessage = "Loading Data....";

                        // Handel return. 
                        if (ctrl.CrmConnectionMgr != null && ctrl.CrmConnectionMgr.CrmSvc != null && ctrl.CrmConnectionMgr.CrmSvc.IsReady)
                        {
                            // Assign local property
                            props.OrgUriActual = ctrl.CrmConnectionMgr.CrmSvc.CrmConnectOrgUriActual.ToString();
                            props.OrgUri = ctrl.CrmConnectionMgr.ConnectedOrgPublishedEndpoints[EndpointType.WebApplication];
                            
                            props.FriendlyName = ctrl.CrmConnectionMgr.CrmSvc.ConnectedOrgFriendlyName;
                            props.AuthenticationProviderType = ctrl.CrmConnectionMgr.CrmSvc.OrganizationServiceProxy.ServiceConfiguration.AuthenticationType.ToString();

                            ClientCredentials credentials = ctrl.CrmConnectionMgr.CrmSvc.OrganizationServiceProxy.ClientCredentials;
                            if (credentials.UserName.UserName != null)
                            {
                                props.UserName = credentials.UserName.UserName;
                                props.Password = credentials.UserName.Password;
                            }
                            else if(credentials.Windows.ClientCredential.UserName !=null)
                            {
                                props.DomainName = credentials.Windows.ClientCredential.Domain;
                                props.UserName = credentials.Windows.ClientCredential.UserName;
                                props.Password = credentials.Windows.ClientCredential.Password;
                            }
                        }
                        
                        // Then generate assembly
                        await Task.Run(() => LoadData());

                        // Set Context class name.
                        props._cxInfo.CustomTypeInfo.CustomTypeName = "Microsoft.Pfe.Xrm.XrmContext";

                        IsLoading = false;
                    }
                    else
                        MessageBox.Show("BadConnect");
                });
            }
        }

        /// <summary>
        /// Login by using CrmLogin and generate an assembly
        /// </summary>
        public RelayCommand ChangeCredentialCommand
        {
            get
            {
                return new RelayCommand(() =>
                {
                    // Establish the Login control
                    CrmLogin ctrl = new CrmLogin();
                    // Wire Event to login response. 
                    ctrl.ConnectionToCrmCompleted += ctrl_ConnectionToCrmCompleted;
                    // Show the dialog. 
                    ctrl.ShowDialog();

                    // Handel return. 
                    if (ctrl.CrmConnectionMgr != null && ctrl.CrmConnectionMgr.CrmSvc != null && ctrl.CrmConnectionMgr.CrmSvc.IsReady)
                    {
                        IsLoading = true;
                        LoadMessage = "Updating Login Credential....";

                        // Handel return. 
                        if (ctrl.CrmConnectionMgr != null && ctrl.CrmConnectionMgr.CrmSvc != null && ctrl.CrmConnectionMgr.CrmSvc.IsReady)
                        {
                            // Assign local property
                            props.OrgUriActual = ctrl.CrmConnectionMgr.CrmSvc.CrmConnectOrgUriActual.ToString();
                            props.OrgUri = ctrl.CrmConnectionMgr.ConnectedOrgPublishedEndpoints[EndpointType.WebApplication];

                            props.FriendlyName = ctrl.CrmConnectionMgr.CrmSvc.ConnectedOrgFriendlyName;
                            props.AuthenticationProviderType = ctrl.CrmConnectionMgr.CrmSvc.OrganizationServiceProxy.ServiceConfiguration.AuthenticationType.ToString();

                            ClientCredentials credentials = ctrl.CrmConnectionMgr.CrmSvc.OrganizationServiceProxy.ClientCredentials;
                            if (credentials.UserName.UserName != null)
                            {
                                props.UserName = credentials.UserName.UserName;
                                props.Password = credentials.UserName.Password;
                            }
                            else if (credentials.Windows.ClientCredential.UserName != null)
                            {
                                props.DomainName = credentials.Windows.ClientCredential.Domain;
                                props.UserName = credentials.Windows.ClientCredential.UserName;
                                props.Password = credentials.Windows.ClientCredential.Password;
                            }

                            IsLoaded = true;
                        }
                                               
                        IsLoading = false;
                    }
                    else
                        MessageBox.Show("BadConnect");
                });
            }
        }

        /// <summary>
        /// Update the assembly with latest Metadata
        /// </summary>
        public RelayCommand ReloadCommand
        {
            get
            {
                return new RelayCommand(async () =>
                {

                    // Handel return. 

                    IsLoading = true;
                    LoadMessage = "Loading Data....";

                    await Task.Run(() => LoadData());

                    IsLoading = false;
                });
            }
        }
        #endregion
    }    
}
