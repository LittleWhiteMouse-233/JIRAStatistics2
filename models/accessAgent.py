from jira import JIRA, JIRAError
import os
from typing import Any
from . import fieldStructure as fieldsS
from . import issueData as issueD
from .JQL import JQLFilter
from .support import exceptions as exc


class JIRALogin:
    # server = "https://rd-mokadisplay.tcl.com/pms/"
    server = "https://idisplayvision.com/jira/"

    @classmethod
    def used_basic(cls, username: str = None, password: str = None):
        if username is None:
            username = input("Username: ")
        if password is None:
            password = input("Password: ")
        return JIRAAgency(JIRA(server=cls.server, basic_auth=(username, password)))

    @classmethod
    def used_token(cls, access_token_resource: str = None):
        if access_token_resource is None:
            access_token_resource = input(
                "access_token string or access_token filepath(input N/n use username&password):")
        if os.path.exists(access_token_resource):
            access_token = open(access_token_resource, 'r').readline()
        else:
            access_token = access_token_resource
        return JIRAAgency(JIRA(server=cls.server, token_auth=access_token))


class JIRAAgency:
    def __init__(self, jira_obj: JIRA):
        self.__jira = jira_obj

    def search_by_jql_filter(self, jql_filter: JQLFilter):
        return self.__jira.search_issues(jql_str=jql_filter.content, startAt=0, maxResults=False)

    def get_project_issue_fields(self, project: fieldsS.Project, issue_type: fieldsS.IssueType):
        return self.__jira.project_issue_fields(project=str(project.id), issue_type=str(issue_type.id),
                                                startAt=0, maxResults=False)

    def get_fields(self):
        return self.__jira.fields()

    def get_single_issue(self, key_or_id: str):
        return self.__jira.issue(key_or_id)

    def create_issue(self, issue_data: dict[str, Any]):
        return self.__jira.create_issue(issue_data)

    def create_issues(self, issues_data: list[dict[str, Any]]):
        return self.__jira.create_issues(issues_data)

    def add_attachment(self, issue_id: str | int, filepath: str):
        self.__jira.add_attachment(issue_id, attachment=filepath)

    def add_comment(self, issue_id: str | int, content: str):
        self.__jira.add_comment(issue_id, body=content)

    def update_comment(self, issue_id: str | int, comment_index: int, content: str):
        comment_list = self.__jira.comments(issue_id)
        comment_list[comment_index].update(body=content)

    def update_latest_comment(self, issue_id: str | int, content: str):
        self.update_comment(issue_id, -1, content)


class JIRAOperator:
    def __init__(self, agency: JIRAAgency):
        self.__agency = agency
        self.__fields = fieldsS.FieldList(self.__agency.get_fields())
        self.__cache = issueD.IssueList()
        self.__num_dict = {
            'call_agency': 0,
            'call_find': 0,
        }

    @property
    def ref_fields(self):
        return self.__fields

    @property
    def call_num_log(self):
        return str(self.__num_dict)

    def add_cache(self, issue_list: list[issueD.Issue]):
        for issue in issue_list:
            self.__cache.append(issue)

    def find_issue_by(self, key_or_id: str):
        self.__num_dict['call_find'] += 1
        cache = self.__cache.self_search_by(key_or_id)
        if cache is not None:
            return cache
        issue_obj = self.__agency.get_single_issue(key_or_id)
        self.__num_dict['call_agency'] += 1
        issue = issueD.Issue.auto_adapt(issue_obj, self.__fields)
        self.__cache.append(issue)
        return issue

    def find_parents(self, issue: issueD.TaskLike):
        if issubclass(type(issue), issueD.Subtask):
            issue: issueD.Subtask
            try:
                parent = self.find_issue_by(issue.parent.key)
            except JIRAError as e:
                raise exc.GetParentFailedError(issue.parent.key, e.text)
        else:
            parent = issue
        if issubclass(type(parent), issueD.Epic):
            epic = parent
            task = None
        else:
            task = parent
            try:
                epic = self.find_issue_by(task.epic_link)
            except JIRAError as e:
                raise exc.GetEpicFailedError(task.epic_link, e.text)
        task: issueD.Task | None
        epic: issueD.Epic
        return task, epic
