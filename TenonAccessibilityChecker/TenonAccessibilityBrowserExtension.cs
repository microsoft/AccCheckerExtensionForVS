/*------------------------------------------- START OF LICENSE -----------------------------------------
HTML Accessibility Checker
Copyright (c) Microsoft Corporation
All rights reserved.
MIT License
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the ""Software""), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
----------------------------------------------- END OF LICENSE ------------------------------------------*/

using System.ComponentModel.Composition;
using System.IO;
using System.Windows;
using Microsoft.VisualStudio.Web.BrowserLink;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;

namespace Microsoft.TenonAccessibilityChecker
{
    [Export(typeof(IBrowserLinkExtensionFactory))]
    public class TenonAccessibilityBrowserExtensionFactory : IBrowserLinkExtensionFactory
    {
        public BrowserLinkExtension CreateExtensionInstance(BrowserLinkConnection connection)
        {
            return new TenonAccessibilityBrowserExtension();
        }

        public string GetScript()
        {
            using (Stream stream = GetType().Assembly.GetManifestResourceStream("Microsoft.TenonAccessibilityChecker.TenonAccessibilityBrowserExtension.js"))
            using (StreamReader reader = new StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }
    }

    public class TenonAccessibilityBrowserExtension : BrowserLinkExtension
    {
        private string connectionURL { get; set; }

        private IVsHierarchy projectHierarchy { get; set; }

        public override void OnConnected(BrowserLinkConnection connection)
        {
            IVsSolution solution = Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(SVsSolution)) as IVsSolution;

            IVsHierarchy hierarchy1 =null;

            if (connection.Project != null)
            {
                int hr = solution.GetProjectOfUniqueName(connection.Project.UniqueName, out hierarchy1);
            }

            projectHierarchy = hierarchy1;
            connectionURL = connection.Url.ToString();
        }

        [BrowserLinkCallback] // This method can be called from JavaScript
        public void SendText(string message)
        {
            if (!TenonAccessibilityCheckerPackage.TenonBrowserLinkSetting)
            {
                TenonModal.InvokedFromBrowser = true;
                TenonModal.RenderedContent = message;

                // close all the open TenonAccessibility dialogs
                CloseAllWindows();

                var dialog = new TenonModal();

                TenonAccessibilityCheckerPackage.ItemFullPath = string.Empty;
                TenonAccessibilityCheckerPackage.Hierarchy = null;
                dialog.Topmost = true;
                TenonAccessibilityCheckerPackage.ItemFullPath = connectionURL;
                TenonAccessibilityCheckerPackage.Hierarchy = projectHierarchy;

                //Show the tenon modal browser dialog
                dialog.ShowDialog();
            }
        }

        /// <summary>
        /// close the browser extension modal window.
        /// </summary>
        private static void CloseAllWindows()
        {
            //ensure only one modal dialog is open at a time for that page.
            for (int intCounter = Application.Current.Windows.Count - 1; intCounter > 0; intCounter--)
            {
                var window = Application.Current.Windows[intCounter];
                if (window != null) window.Close();
            }
        }

    }
}