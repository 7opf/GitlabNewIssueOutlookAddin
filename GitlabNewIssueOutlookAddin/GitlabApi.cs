using RestSharp;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GitlabNewIssueOutlookAddin {
    public class GitlabApi {

        RestClient rc;

        public GitlabApi() {
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
            rc = new RestClient("https://gitlab.eps.surrey.ac.uk/api/v4");
            rc.AddDefaultHeader("PRIVATE-TOKEN", "");
        }

        public GitlabIssue newIssue(int projectId, GitlabNewIssue newIssue) {
            RestRequest req = new RestRequest("projects/{projectId}/issues");
            req.AddParameter("projectId", projectId, ParameterType.UrlSegment);
            req.AddJsonBody(@newIssue);

            try {
                var response = rc.Post<GitlabIssue>(req);
                Debug.WriteLine(response.Content);

                if (response.StatusCode != System.Net.HttpStatusCode.Created) {
                    MessageBox.Show("Failed to create issue: " + response.Content);
                    return null;
                }

                return response.Data;
            } catch (Exception e) {
                Debug.WriteLine(e);
                MessageBox.Show(e.Message);
                return null;
            }

        }

        public List<GitlabSimpleProject> getProjects() {
            RestRequest req = new RestRequest("projects");
            req.AddQueryParameter("order_by", "name");
            req.AddQueryParameter("sort", "asc");
            req.AddQueryParameter("simple", "true");
            req.AddQueryParameter("search", "tihm");
            req.AddQueryParameter("membership", "true");
            req.AddQueryParameter("with_issues_enabled", "true");

            try {
                var response = rc.Get<List<GitlabSimpleProject>>(req);
                Debug.WriteLine(response.Content);

                if (response.StatusCode != System.Net.HttpStatusCode.OK) {
                    MessageBox.Show("Failed to fetch Gitlab project list: " + response.Content);
                    return null;
                }

                return response.Data;
            } catch (Exception e) {
                Debug.WriteLine(e);
                MessageBox.Show(e.Message);
                return null;
            }
        }

    }

    public class GitlabIssue {
        // https://docs.gitlab.com/ee/api/issues.html#new-issue
        private int _id;
        public int id {
            get {
                return _id;
            }
            set {
                _id = value;
            }
        }
        private String _web_url;
        public String web_url {
            get {
                return _web_url;
            }
            set {
                _web_url = value;
            }
        }
        public GitlabIssue() { }
    }

    public class GitlabNewIssue {
        // https://docs.gitlab.com/ee/api/issues.html#new-issue
        private String _title;
        public String title {
            get {
                return _title;
            }
            set {
                _title = value;
            }
        }
        private String _description;
        public String description {
            get {
                return _description;
            }
            set {
                _description = value;
            }
        }
        private Boolean _confidential;
        public Boolean confidential {
            get {
                return _confidential;
            }
            set {
                _confidential = value;
            }
        }
        private List<int> _assignee_ids;
        public List<int> assignee_ids {
            get {
                return _assignee_ids;
            }
            set {
                _assignee_ids = value;
            }
        }
        private int _milestone_id;
        public int milestone_id {
            get {
                return _milestone_id;
            }
            set {
                _milestone_id = value;
            }
        }
        private String _labels;
        public String labels {
            get {
                return _labels;
            }
            set {
                _labels = value;
            }
        }
        private String _created_at;
        public String created_at {
            get {
                return _created_at;
            }
            set {
                _created_at = value;
            }
        }
        private String _due_date;
        public String due_date {
            get {
                return _due_date;
            }
            set {
                _due_date = value;
            }
        }
        private int _merge_request_to_resolve_discussions_of;
        public int merge_request_to_resolve_discussions_of {
            get {
                return _merge_request_to_resolve_discussions_of;
            }
            set {
                _merge_request_to_resolve_discussions_of = value;
            }
        }
        private String _discussion_to_resolve;
        public String discussion_to_resolve {
            get {
                return _discussion_to_resolve;
            }
            set {
                _discussion_to_resolve = value;
            }
        }
        private int _weight;
        public int weight {
            get {
                return _weight;
            }
            set {
                _weight = value;
            }
        }

        public GitlabNewIssue() { }
    }

    public class GitlabSimpleProject {
        // https://docs.gitlab.com/ee/api/projects.html#list-all-projects
        private int _id;
        public int id {
            get {
                return _id;
            }
            set {
                _id = value;
            }
        }
        private String _description;
        public String description {
            get {
                return _description;
            }
            set {
                _description = value;
            }
        }
        private String _default_branch;
        public String default_branch {
            get {
                return _default_branch;
            }
            set {
                _default_branch = value;
            }
        }
        private String _ssh_url_to_repo;
        public String ssh_url_to_repo {
            get {
                return _ssh_url_to_repo;
            }
            set {
                _ssh_url_to_repo = value;
            }
        }
        private String _http_url_to_repo;
        public String http_url_to_repo {
            get {
                return _http_url_to_repo;
            }
            set {
                _http_url_to_repo = value;
            }
        }
        private String _web_url;
        public String web_url {
            get {
                return _web_url;
            }
            set {
                _web_url = value;
            }
        }
        private String _readme_url;
        public String readme_url {
            get {
                return _readme_url;
            }
            set {
                _readme_url = value;
            }
        }
        private List<String> _tag_list;
        public List<String> tag_list {
            get {
                return _tag_list;
            }
            set {
                _tag_list = value;
            }
        }
        private String _name;
        public String name {
            get {
                return _name;
            }
            set {
                _name = value;
            }
        }
        private String _name_with_namespace;
        public String name_with_namespace {
            get {
                return _name_with_namespace;
            }
            set {
                _name_with_namespace = value;
            }
        }
        private String _path;
        public String path {
            get {
                return _path;
            }
            set {
                _path = value;
            }
        }
        private String _path_with_namespace;
        public String path_with_namespace {
            get {
                return _path_with_namespace;
            }
            set {
                _path_with_namespace = value;
            }
        }
        private String _created_at;
        public String created_at {
            get {
                return _created_at;
            }
            set {
                _created_at = value;
            }
        }
        private String _last_activity_at;
        public String last_activity_at {
            get {
                return _last_activity_at;
            }
            set {
                _last_activity_at = value;
            }
        }
        private int _forks_count;
        public int forks_count {
            get {
                return _forks_count;
            }
            set {
                _forks_count = value;
            }
        }
        private String _avatar_url;
        public String avatar_url {
            get {
                return _avatar_url;
            }
            set {
                _avatar_url = value;
            }
        }
        private int _star_count;
        public int star_count {
            get {
                return _star_count;
            }
            set {
                _star_count = value;
            }
        }

        public GitlabSimpleProject() { }
    }
}
