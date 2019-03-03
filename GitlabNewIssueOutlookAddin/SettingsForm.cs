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
            this.textBox1.Text = this.gitlabApi.Url;
            this.textBox2.Text = this.gitlabApi.Token;
            foreach (String[] p in this.gitlabApi.Parameters) {
                this.checkBox1.Checked = false;
                this.checkBox2.Checked = false;
                this.checkBox3.Checked = false;
                switch (p[0]) {
                    case "search":
                        this.textBox3.Text = p[1];
                        break;
                    case "membership":
                        this.checkBox1.Checked = true;
                        break;
                    case "owned":
                        this.checkBox2.Checked = true;
                        break;
                    case "starred":
                        this.checkBox3.Checked = true;
                        break;
                }
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            Process.Start("https://docs.gitlab.com/ee/user/profile/personal_access_tokens.html");
        }

        private void button1_Click(object sender, EventArgs e) {
            // ok
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e) {
            // cancel
            this.Close();
        }

        private void SettingsForm_FormClosing(object sender, FormClosingEventArgs e) {
            if (this.DialogResult != DialogResult.OK) {
                return;
            }

            // update the api after clicking ok
            try {
                if (this.textBox1.Text != null) {
                    this.gitlabApi.Url = this.textBox1.Text;
                }

                if (this.textBox2.Text != null) {
                    this.gitlabApi.Token = this.textBox2.Text;
                }

                List<String[]> p = new List<String[]> { };
                if (this.textBox3.Text != null) {
                    p.Add(new string[] { "search", this.textBox3.Text });
                }
                if (this.checkBox1.Checked) {
                    p.Add(new string[] { "membership", "true" });
                }
                if (this.checkBox2.Checked) {
                    p.Add(new string[] { "owned", "true" });
                }
                if (this.checkBox3.Checked) {
                    p.Add(new string[] { "starred", "true" });
                }
                this.gitlabApi.Parameters = p.ToArray();
            } catch (Exception err) {
                e.Cancel = true;
                MessageBox.Show(err.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
    }
}
