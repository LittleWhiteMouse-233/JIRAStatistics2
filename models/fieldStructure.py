import jira.resources as jira_res
from typing import Any
from . import utils


class Field:
    def __init__(self, field_obj):
        if type(field_obj) is dict:
            self.id = field_obj['id']
            self.name = field_obj['name']
        else:
            self.id = field_obj.fieldId
            self.name = field_obj.name


class FieldList(list[Field]):
    def __init__(self, field_obj_list: list[dict[str, Any]]):
        super().__init__()
        for field_obj in field_obj_list:
            self.append(Field(field_obj))

    def field_name2id(self, field_name: str):
        for field in self:
            if field.name == field_name:
                return field.id
        raise ValueError("The field_name(%s) is not found." % field_name)


class Project:
    def __init__(self, project_obj: jira_res.Project):
        self.key = project_obj.key
        self.name = project_obj.name
        self.id = project_obj.id


class IssueType:
    def __init__(self, issue_type_obj: jira_res.IssueType):
        self.name = issue_type_obj.name
        self.id = issue_type_obj.id
        self.isSubtask = issue_type_obj.subtask


class WorkflowStatus:
    def __init__(self, status_obj: jira_res.Status):
        self.name = status_obj.name
        self.id = status_obj.id
        self.statusCategory = status_obj.statusCategory.id


class Priority:
    def __init__(self, priority_obj: jira_res.Priority):
        self.name = priority_obj.name
        self.id = priority_obj.id


class User:
    def __init__(self, user_obj: jira_res.User):
        self.displayName = user_obj.displayName
        self.key = user_obj.key
        self.accountName = user_obj.name
        self.emailAddress = user_obj.emailAddress


class Comment:
    def __init__(self, comment_obj: jira_res.Comment):
        self.body = utils.clean_string(comment_obj.body)
        self.created_author = User(comment_obj.author)
        self.created_timestring = comment_obj.created
        self.updated_author = User(comment_obj.updateAuthor)
        self.updated_timestring = comment_obj.updated


class Worklog:
    def __init__(self, worklog_obj: jira_res.Worklog):
        self.created_author = User(worklog_obj.author)
        self.created_timestring = worklog_obj.created
        self.comment = worklog_obj.comment
        self.id = worklog_obj.id
        self.issueId = worklog_obj.issueId
        self.started_timestring = worklog_obj.started
        self.timeSpent = worklog_obj.timeSpent
        self.timeSpentSeconds = worklog_obj.timeSpentSeconds
        self.updated_author = User(worklog_obj.updateAuthor)
        self.updated_timestring = worklog_obj.updated


class OptionValue:
    def __init__(self, customfield_obj: jira_res.CustomFieldOption):
        self.id = customfield_obj.id
        self.value = customfield_obj.value


class MultOptionValue:
    def __init__(self, customfield_obj: jira_res.CustomFieldOption):
        self.parent = OptionValue(customfield_obj)
        self.child = OptionValue(customfield_obj.child)


class Component:
    def __init__(self, components_obj: jira_res.Component):
        self.id = components_obj.id
        self.name = components_obj.name
