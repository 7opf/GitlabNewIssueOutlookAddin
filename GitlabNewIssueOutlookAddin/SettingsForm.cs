using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GitlabNewIssueOutlookAddin {
    public partial class SettingsForm : Form {

        private GitlabApi gitlabApi;

        public SettingsForm(GitlabApi gitlabApi) {
            InitializeComponent();
            this.gitlabApi = gitlabApi;
            this.textBox1.Text = Properties.Settings.Default.Url;
            this.textBox2.Text = Properties.Settings.Default.Token;
            this.textBox3.Text = Properties.Settings.Default.Search;
            this.checkBox1.Checked = Properties.Settings.Default.Membership;
            this.checkBox2.Checked = Properties.Settings.Default.Owned;
            this.checkBox3.Checked = Properties.Settings.Default.Starred;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            Process.Start("https://docs.gitlab.com/ee/user/profile/personal_access_tokens.html");
        }

        private void button1_Click(object sender, EventArgs e) {
            // ok
            this.DialogResult = DialogResult.OK;
            Properties.Settings prevSettings = new Properties.Settings();
            prevSettings = Properties.Settings.Default;

            Properties.Settings.Default.Url = this.textBox1.Text;
            Properties.Settings.Default.Token = this.textBox2.Text;
            Properties.Settings.Default.Search = this.textBox3.Text;
            Properties.Settings.Default.Membership = this.checkBox1.Checked;
            Properties.Settings.Default.Owned = this.checkBox2.Checked;
            Properties.Settings.Default.Starred = this.checkBox3.Checked;

            try {
                this.gitlabApi.Configure();
                if (this.gitlabApi.Configured) {
                    this.gitlabApi.fetchProjects();
                    // only save settings upon successful configuration
                    Properties.Settings.Default.Save();
                }
                this.Close();
            } catch (Exception err) {
                Debug.WriteLine(err);
                MessageBox.Show(err.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e) {
            // cancel
            Properties.Settings.Default.Reload(); // reset in-memory settings
            this.Close();
        }
    }
}
