using System;
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

        public GitlabNewIssueRibbon() { }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID) {
            return GetResourceText("GitlabNewIssueOutlookAddin.GitlabNewIssueRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI) {
            this.ribbon = ribbonUI;
            this.gitlabApi = new GitlabApi();
            try {
                this.gitlabApi.Configure();
                if (this.gitlabApi.Configured) {
                    this.gitlabApi.fetchProjects();
                }
            } catch (Exception err) {
                // ignore any errors onLoad
                Debug.WriteLine(err);
            }
        }

        public void OpenSettings(Office.IRibbonControl control) {
            SettingsForm form = new SettingsForm(this.gitlabApi);
            form.Show();
        }

        public String PopulateMenu(Office.IRibbonControl control) {
            return GetMenuXML(this.gitlabApi.Projects);
        }

        public void SubmitIssue(Office.IRibbonControl control) {
            Outlook.MailItem mail = null;

            if (control.Context is Outlook.Selection) {
                Outlook.Selection sel = control.Context as Outlook.Selection;
                mail = sel[1];
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

            try {
                GitlabIssue createdIssue = this.gitlabApi.newIssue(Int32.Parse(control.Tag), issue);
                DialogResult result = MessageBox.Show($"View on Gitlab?", "Issue Created", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes) {
                    Process.Start(createdIssue.web_url);
                }
            } catch (Exception err) {
                Debug.WriteLine(err);
                MessageBox.Show(err.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
            if (projects == null) {
                return @"<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui""><button id=""button1"" label=""No projects available"" enabled=""false"" /></menu>";
            }
            return $@"<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">{String.Join("", projects.Select(GetButtonXML))}</menu>";
        }

        public String GetButtonXML(GitlabSimpleProject project) {
            return $@"<button id=""project{project.id}"" label=""{project.name_with_namespace}"" tag=""{project.id}"" onAction=""SubmitIssue"" />";
        }

        #endregion
    }
}
