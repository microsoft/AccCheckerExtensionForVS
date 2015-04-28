/*------------------------------------------- START OF LICENSE -----------------------------------------
HTML Accessibility Checker
Copyright (c) Microsoft Corporation
All rights reserved.
MIT License
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the ""Software""), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
----------------------------------------------- END OF LICENSE ------------------------------------------*/


using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Newtonsoft.Json;
using EnvDTE80;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.ComponentModel;
using System.Security;
using System.Windows;

namespace Microsoft.TenonAccessibilityChecker
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell.
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    // This attribute is used to register the information needed to show this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]

    //This is needed to load the package automatically on install.
    [ProvideAutoLoad("{f1536ef8-92ec-443c-9ed7-fdadf150da82}")]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.GuidTenonAccessibilityCheckerPkgString)]
    public sealed class TenonAccessibilityCheckerPackage : Package
    {
        /////////////////////////////////////////////////////////////////////////////
        // Overridden Package Implementation
        #region Package Members
        public static IVsHierarchy Hierarchy { get; set; }
        public static string ItemFullPath { get; set; }

        public static string TenonStatusCode { get; set; }
        public static string TenonErrorMessage { get; set; }

        public static bool TenonBrowserLinkSetting { get; set; }

        private  DTE2 m_applicationObject = null;
        private SolutionEvents solutionEvents;

        public  DTE2 ApplicationObject
        {
            get
            {
                if (m_applicationObject == null)
                {
                    // Get an instance of the currently running Visual Studio IDE
                    DTE dte = (DTE)GetService(typeof(DTE));
                    m_applicationObject = dte as DTE2;
                }
                return m_applicationObject;
            }
        }

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            base.Initialize();
            TaskManager.Initialize(this);

            // Add our command handlers for menu (commands must exist in the .vsct file)
            var mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null == mcs) return;

            // Create the command for the menu item.
            var menuCommandId = new CommandID(GuidList.GuidTenonAccessibilityCheckerCmdSet, (int)PkgCmdId.CmdidMyTenonCommand);
            var menuItem = new OleMenuCommand(MenuItemCallback, menuCommandId);

            var menuCommandId1 = new CommandID(GuidList.GuidTenonAccessibilityCheckerCmdSet, (int)PkgCmdId.CmdidMyBrowserExtension);
            var menuItem1 = new OleMenuCommand(MenuItemCallback1, menuCommandId1);

            menuItem.BeforeQueryStatus += menuCommand_BeforeQueryStatus;
            mcs.AddCommand(menuItem);

            menuItem1.BeforeQueryStatus += menuCommand_BeforeQueryStatus1;
            mcs.AddCommand(menuItem1);

            solutionEvents = ApplicationObject.Events.SolutionEvents;
            solutionEvents.AfterClosing += new _dispSolutionEvents_AfterClosingEventHandler(SolutionAfterClosing);
        }
        #endregion

        private void MenuItemCallback1(object sender, EventArgs e)
        {
            TenonBrowserLinkSetting = !TenonBrowserLinkSetting;
        }
        public void SolutionAfterClosing()
        {
            TaskManager.ClearErrors();
        }

        private void menuCommand_BeforeQueryStatus1(object sender, EventArgs e)
        {
            var menuCommand = sender as OleMenuCommand;

            if (menuCommand != null)
            {
                menuCommand.Checked = !TenonBrowserLinkSetting;
            }
        }

        /// <summary>
        /// Validation to display the "check accessibility with tenon" for certain file type extensions.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuCommand_BeforeQueryStatus(object sender, EventArgs e)
        {
            var menuCommand = sender as OleMenuCommand;

            if (menuCommand != null)
            {
                IntPtr hierarchyPtr, selectionContainerPtr;
                uint projectItemId;
                IVsMultiItemSelect mis;
                var monitorSelection = (IVsMonitorSelection)GetGlobalService(typeof(SVsShellMonitorSelection));
                monitorSelection.GetCurrentSelection(out hierarchyPtr, out projectItemId, out mis, out selectionContainerPtr);

                var hierarchy = Marshal.GetTypedObjectForIUnknown(hierarchyPtr, typeof(IVsHierarchy)) as IVsHierarchy;
                if (hierarchy != null)
                {
                    object value;
                    hierarchy.GetProperty(projectItemId, (int)__VSHPROPID.VSHPROPID_Name, out value);

                    string extension = value.ToString();

                    if (value != null &&
                        (extension.EndsWith(".cshtml", StringComparison.OrdinalIgnoreCase) ||
                         extension.EndsWith(".xhtml", StringComparison.OrdinalIgnoreCase) ||
                         extension.EndsWith(".html", StringComparison.OrdinalIgnoreCase) ||
                         extension.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase) ||
                         extension.EndsWith(".ascx", StringComparison.OrdinalIgnoreCase) ||
                         extension.EndsWith(".asp", StringComparison.OrdinalIgnoreCase) ||
                         extension.EndsWith(".htm", StringComparison.OrdinalIgnoreCase)))
                    {
                        menuCommand.Visible = true;
                    }
                    else
                    {
                        menuCommand.Visible = false;
                    }
                }
            }
        }

        /// <summary>
        /// Check if an single item is selected in the project.
        /// </summary>
        /// <param name="hierarchy"></param>
        /// <param name="itemId"></param>
        /// <returns></returns>
        public static bool IsSingleProjectItemSelection(out IVsHierarchy hierarchy, out uint itemId)
        {
            hierarchy = null;
            itemId = VSConstants.VSITEMID_NIL;

            var monitorSelection = GetGlobalService(typeof(SVsShellMonitorSelection)) as IVsMonitorSelection;
            var solution = GetGlobalService(typeof(SVsSolution)) as IVsSolution;
            if (monitorSelection == null || solution == null)
            {
                return false;
            }

            var hierarchyPtr = IntPtr.Zero;
            var selectionContainerPtr = IntPtr.Zero;

            try
            {
                IVsMultiItemSelect multiItemSelect;
                var hr = monitorSelection.GetCurrentSelection(out hierarchyPtr, out itemId, out multiItemSelect, out selectionContainerPtr);

                if (ErrorHandler.Failed(hr) || hierarchyPtr == IntPtr.Zero || itemId == VSConstants.VSITEMID_NIL)
                {
                    return false;
                }

                // multiple items are selected
                if (multiItemSelect != null) return false;

                // there is a hierarchy root node selected, thus it is not a single item inside a project
                if (itemId == VSConstants.VSITEMID_ROOT) return false;

                hierarchy = Marshal.GetObjectForIUnknown(hierarchyPtr) as IVsHierarchy;
                if (hierarchy == null) return false;

                Guid guidProjectId;

                return !ErrorHandler.Failed(solution.GetGuidOfProject(hierarchy, out guidProjectId));
            }
            finally
            {
                if (selectionContainerPtr != IntPtr.Zero)
                {
                    Marshal.Release(selectionContainerPtr);
                }

                if (hierarchyPtr != IntPtr.Zero)
                {
                    Marshal.Release(hierarchyPtr);
                }
            }
        }

        private static ProjectItem GetProjectItemFromHierarchy(IVsHierarchy pHierarchy, uint itemId)
        {
            object propertyValue;
            ErrorHandler.ThrowOnFailure(pHierarchy.GetProperty(itemId, (int)__VSHPROPID.VSHPROPID_ExtObject, out propertyValue));

            var projectItem = propertyValue as ProjectItem;
            if (projectItem == null) return null;

            return projectItem;
        }



        /// <summary>
        /// This function is the callback used to execute a command when the a menu item is clicked.
        /// See the Initialize method to see how the menu item is associated to this function using
        /// the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            IVsHierarchy hierarchy;
            uint itemid;

            if (!IsSingleProjectItemSelection(out hierarchy, out itemid)) return;

            // ReSharper disable once SuspiciousTypeConversion.Global
            var vsProject = (IVsProject)hierarchy;

            string projectFullPath;
            if (vsProject != null && ErrorHandler.Failed(vsProject.GetMkDocument(VSConstants.VSITEMID_ROOT, out projectFullPath))) return;

            // get the name of the item
            string itemFullPath = null;
            if (vsProject != null && ErrorHandler.Failed(vsProject.GetMkDocument(itemid, out itemFullPath))) return;

            var selectedProjectItem = GetProjectItemFromHierarchy(hierarchy, itemid);

            if (selectedProjectItem == null) return;

            if (itemFullPath != null)
            {
                ItemFullPath = itemFullPath;
                Hierarchy = hierarchy;
                TenonModal.InvokedFromBrowser = false;
                TenonModal customdialog = new TenonModal();
                customdialog.ShowModal();
            }
        }

        /// <summary>
        /// parse the Tenon API response.
        /// </summary>
        /// <param name="JsonString">Response from Tenon API</param>
        /// <returns>ErrorCollection List</returns>
        public static IList<ErrorResultSet> ParseJson(string jsonOutput)
        {
            //parse string
            dynamic dynObj = JsonConvert.DeserializeObject(jsonOutput);

            if (dynObj != null)
            {
                TenonStatusCode = (string)dynObj.status;

                if ((string)dynObj.status != contents.TenonApiResponseSuccess)
                {
                    TenonErrorMessage = (string)dynObj.message;
                    return null;
                };

                return ProcessResults(dynObj);
            }

            return null;
        }

        /// <summary>
        /// Process the json result output from the tenon api
        /// </summary>
        /// <param name="dynObj"></param>
        /// <returns></returns>
        private static IList<ErrorResultSet> ProcessResults(dynamic dynObj)
        {
            var eCollectionResult = new List<ErrorResultSet>();

            foreach (var data1 in dynObj.resultSet)
            {
                var eCollection = new ErrorResultSet
                {
                    Certainty = data1.certainty,
                    ErrorTitle = (string)data1.errorTitle,
                    ErrorDescription = WebUtility.HtmlDecode((string)data1.errorDescription),
                    Column = !string.IsNullOrEmpty(data1.position.ToString()) ? data1.position.column.ToString() : string.Empty,
                    Line = !string.IsNullOrEmpty(data1.position.ToString()) ? data1.position.line.ToString() : string.Empty,
                    Referencelink = data1.@ref
                };

                eCollectionResult.Add(eCollection);
            }

            TenonErrorMessage = string.Format(contents.TenonAPISuccess, dynObj.resultSummary.issues.totalErrors.ToString(), dynObj.resultSummary.issues.totalWarnings.ToString());

            return eCollectionResult;
        }

        /// <summary>
        /// call the Tenon API for the provided inputs.
        /// </summary>
        /// <param name="key"></param>
        /// <param name="source"></param>
        /// <param name="certainty"></param>
        /// <param name="level"></param>
        /// <returns></returns>
        public static string GetJson(string key, string source, string certainty, string level)
        {
            //replace & to amp;
            var strsource = !string.IsNullOrEmpty(source) ? source.Replace("&", "%26") : string.Empty;

            var responseFromServer = string.Empty;
            HttpWebRequest req = null;

            try
            {
                req = (HttpWebRequest)WebRequest.Create(new Uri(contents.TenonService));

                ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;

                req.Method = "POST";

                var encoding = new ASCIIEncoding();

                var postData = "key=" + key;
                postData += ("&src=" + strsource);
                postData += ("&inline=" + false);
                postData += ("&certainty=" + (certainty == "Errors and Warnings" ? 0 : 80));
                postData += ("&level=" + level);
                postData += ("&fragment=" + 0);
                postData += ("&importance=" + 1);
                postData += ("&priority=" + 0);
                postData += ("&store=" + 0);
                postData += ("&ref=" + 1);
                postData += ("&appID=" + contents.TenonAppId);

                var data = encoding.GetBytes(postData);

                req.ContentType = "application/x-www-form-urlencoded";

                req.ContentLength = data.Length;

                // Get the request stream.
                var dataStream = req.GetRequestStream();
                // Write the data to the request stream.
                dataStream.Write(data, 0, data.Length);
                // Close the Stream object.
                dataStream.Close();
                // Get the response.
                var response = req.GetResponse();

                // Get the stream containing content returned by the server.
                dataStream = response.GetResponseStream();
                // Open the stream using a StreamReader for easy access.
                if (dataStream != null)
                {
                    var reader = new StreamReader(dataStream);
                    // Read the content.
                    responseFromServer = reader.ReadToEnd();

                    // Clean up the streams.
                    reader.Close();
                    //dataStream.Close();
                    response.Close();

                    return responseFromServer;
                }

                //return responseFromServer;
            }
            catch (WebException ex)
            {
                if (ex.Status == WebExceptionStatus.Timeout)
                {
                    TenonErrorMessage = ex.Message.ToString();
                    TenonStatusCode = "999";
                }
                else
                {
                    TenonErrorMessage = ex.Message.ToString();
                    TenonStatusCode = "1000";
                }
                throw;
            }
            catch (Exception ex)
            {
                TenonErrorMessage = ex.Message.ToString();
                TenonStatusCode = "2000";
                throw;
            }
            return responseFromServer;
        }

        /// <summary>
        /// Validating the tenon api SSL certificate
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="certificate"></param>
        /// <param name="chain"></param>
        /// <param name="sslPolicyErrors"></param>
        /// <returns></returns>
        private static bool CertificateValidationCallBack( object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            // If the certificate is a valid, signed certificate, return true.
            if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
            {
                return true;
            }

            // If there are errors in the certificate chain, look at each error to determine the cause.
            if ((sslPolicyErrors & System.Net.Security.SslPolicyErrors.RemoteCertificateChainErrors) != 0)
            {
                if (chain != null && chain.ChainStatus != null)
                {
                    foreach (System.Security.Cryptography.X509Certificates.X509ChainStatus status in chain.ChainStatus)
                    {
                        if ((certificate.Subject == certificate.Issuer) &&
                           (status.Status == System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.UntrustedRoot))
                        {
                            // Self-signed certificates with an untrusted root are valid.
                            continue;
                        }
                        else
                        {
                            if (status.Status != System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.NoError)
                            {
                                // If there are any other errors in the certificate chain, the certificate is invalid,
                                // so the method returns false.
                                return false;
                            }
                        }
                    }
                }

                // When processing reaches this line, the only errors in the certificate chain are
                // untrusted root errors for self-signed certificates. These certificates are valid
                // for default Exchange server installations, so return true.
                return true;
            }
            else
            {
                // In all other cases, return false.
                return false;
            }
        }


        /// <summary>
        /// Write errors to visual  studio error list.
        /// </summary>
        /// <param name="itemFullPath"></param>
        /// <param name="hierarchy"></param>
        /// <param name="errors"></param>
        public static void WriteErrors(string itemFullPath, IVsHierarchy hierarchy, IList<ErrorResultSet> errors)
        {
            TaskManager.ClearErrors();

            var itemFileName = itemFullPath;

            if (errors != null && errors.Count > 0)
            {
                foreach (var eCollection in errors)
                {
                    //As per Tenon API, certainty greater than equal to 80 is classified as an error.
                    if (eCollection.Certainty >= 80)
                        TaskManager.AddError(eCollection, itemFileName, hierarchy);
                    else
                        TaskManager.AddWarning(eCollection, itemFileName, hierarchy);
                }
            }

            //View ErrorList
            TaskManager.ShowErrorList();
        }

    }
}
