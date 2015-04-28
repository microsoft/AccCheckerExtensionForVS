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
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Globalization;

namespace Microsoft.TenonAccessibilityChecker
{
    /// <summary>
    /// Helper class to write error to visual studio error list.
    /// </summary>
    public static class TaskManager
    {
        private static ErrorListProvider ErrorListProvider;

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="serviceProvider"></param>
        public static void Initialize(IServiceProvider serviceProvider)
        {
           ErrorListProvider = new ErrorListProvider(serviceProvider);
        }

        /// <summary>
        ///Clear the errors in the Visual Studio error List
        /// </summary>
        public static void ClearErrors()
        {
            ErrorListProvider.Tasks.Clear();
        }

        /// <summary>
        /// Focus the visual studio error list
        /// </summary>
        public static void ShowErrorList()
        {
            ErrorListProvider.BringToFront();
        }

        /// <summary>
        /// Add error helper method
        /// </summary>
        /// <param name="message"></param>
        /// <param name="fileName"></param>
        /// <param name="hierarchyItem"></param>
        public static void AddError(ErrorResultSet message, string fileName,IVsHierarchy hierarchyItem)
        {
            AddTask(message, TaskErrorCategory.Error, fileName, hierarchyItem);
        }

        /// <summary>
        /// Add Warning helper method
        /// </summary>
        /// <param name="message"></param>
        /// <param name="fileName"></param>
        /// <param name="hierarchyItem"></param>
        public static void AddWarning(ErrorResultSet message, string fileName,IVsHierarchy hierarchyItem)
        {
            AddTask(message, TaskErrorCategory.Warning, fileName, hierarchyItem);
        }

        /// <summary>
        /// This Method writes errors to visual studio
        /// </summary>
        /// <param name="message"></param>
        /// <param name="category"></param>
        /// <param name="filedetails"></param>
        /// <param name="hierarchyItem"></param>
        private static void AddTask(ErrorResultSet message, TaskErrorCategory category, string filedetails, IVsHierarchy hierarchyItem)
        {
            ErrorListProvider.Tasks.Add(new ErrorTask
            {
                Category = TaskCategory.User,
                ErrorCategory = category,
                Text = message.ErrorDescription + " (" +message.Referencelink +")",
                HierarchyItem = hierarchyItem,
                Column = string.IsNullOrEmpty(message.Column) ? 0 : Int32.Parse(message.Column, CultureInfo.InvariantCulture),
                Line = string.IsNullOrEmpty(message.Line) ? 0 : Int32.Parse(message.Line, CultureInfo.InvariantCulture),
                Document = filedetails
            });
        }
    }
}
