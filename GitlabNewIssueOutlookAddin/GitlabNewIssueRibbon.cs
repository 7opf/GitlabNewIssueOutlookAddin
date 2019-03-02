﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.
//    https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/aa722523(v=office.12)#how-can-i-determine-the-correct-signatures-for-each-callback-procedure

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace GitlabNewIssueOutlookAddin {
    [ComVisible(true)]
    public class GitlabNewIssueRibbon : Office.IRibbonExtensibility {
        private Office.IRibbonUI ribbon;
        private GitlabApi gitlabApi;
        private GitlabSimpleProject[] projects;

        public GitlabNewIssueRibbon() {

        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID) {
            return GetResourceText("GitlabNewIssueOutlookAddin.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI) {
            this.ribbon = ribbonUI;
            this.gitlabApi = new GitlabApi();
            this.projects = this.gitlabApi.getProjects().ToArray();
        }

        public void OpenSettings(Office.IRibbonControl control) {

        }

        public String PopulateMenu(Office.IRibbonControl control) {
            return GetMenuXML(this.projects);
        }

        public void SubmitIssue(Office.IRibbonControl control) {        
            Debug.WriteLine($"Clicked {control.Id}: {control.Tag}");
            Outlook.MailItem mail = null;

            if (control.Context is Outlook.Selection) {
                Outlook.Selection sel = control.Context as Outlook.Selection;
                mail = sel[0];
            }

            if (control.Context is Outlook.MailItem) {
                mail = control.Context as Outlook.MailItem;
            }

            if (mail == null) {
                MessageBox.Show("Please select an email.");
                return;
            }

            GitlabNewIssue issue = new GitlabNewIssue {
                title = mail.Subject,
                description = mail.Body,
                labels = "To Do"
            };

            this.gitlabApi.newIssue(Int32.Parse(control.Id), issue);
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName) {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i) {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0) {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i]))) {
                        if (resourceReader != null) {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        public String GetMenuXML(GitlabSimpleProject[] projects) {
            return $@"<menu description=""Create an issue from this email on Gitlab"" id=""GitlabIssueContextMenu"" label=""Create Gitlab Issue"">{projects.Select(GetButtonXML)}</menu>";
        }

        public String GetButtonXML(GitlabSimpleProject project) {
            return $@"<button id=""{project.id}"" label=""{project.name}"" tag=""{project.name}"" onAction=""SubmitIssue"" />";
        }

        #endregion
    }
}