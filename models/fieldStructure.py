import jira.resources as jira_res
from typing import Any
from dataclasses import dataclass
from .support import utils


@dataclass(slots=True, frozen=True)
class Field:
    id: str
    name: str

    @classmethod
    def init_obj(cls, field_obj: jira_res.Field):
        return cls(
            id=field_obj.fieldId,
            name=field_obj.name,
        )


class FieldList(list[Field]):
    def __init__(self, field_obj_list: list[dict[str, Any]]):
        super().__init__()
        for field_obj in field_obj_list:
            self.append(Field(field_obj['id'], field_obj['name']))

    def field_name2id(self, field_name: str):
        for field in self:
            if field.name == field_name:
                return field.id
        raise ValueError("The field_name(%s) is not found." % field_name)


@dataclass(slots=True, frozen=True)
class Project:
    key: str
    name: str
    id: str

    @classmethod
    def init_obj(cls, project_obj: jira_res.Project):
        return cls(
            key=project_obj.key,
            name=project_obj.name,
            id=project_obj.id,
        )


@dataclass(slots=True, frozen=True)
class IssueType:
    name: str
    id: str
    isSubtask: bool

    @classmethod
    def init_obj(cls, issue_type_obj: jira_res.IssueType):
        return cls(
            name=issue_type_obj.name,
            id=issue_type_obj.id,
            isSubtask=issue_type_obj.subtask,
        )


@dataclass(slots=True, frozen=True)
class WorkflowStatus:
    name: str
    id: str
    statusCategory: str

    @classmethod
    def init_obj(cls, status_obj: jira_res.Status):
        return cls(
            name=status_obj.name,
            id=status_obj.id,
            statusCategory=status_obj.statusCategory.id,
        )


@dataclass(slots=True, frozen=True)
class Priority:
    name: str
    id: str

    @classmethod
    def init_obj(cls, priority_obj: jira_res.Priority):
        return cls(
            name=priority_obj.name,
            id=priority_obj.id,
        )


@dataclass(slots=True, frozen=True)
class User:
    displayName: str
    key: str
    accountName: str
    emailAddress: str

    @classmethod
    def init_obj(cls, user_obj: jira_res.User):
        return cls(
            displayName=user_obj.displayName,
            key=user_obj.key,
            accountName=user_obj.name,
            emailAddress=user_obj.emailAddress,
        )

    @classmethod
    def init_default(cls):
        return cls(
            displayName='auto',
            key='auto',
            accountName='auto',
            emailAddress='',
        )

    @classmethod
    def init_group_item(cls, member_dict:dict[str, ...]):
        return cls(
            displayName=member_dict['fullname'],
            key=member_dict[''],
            accountName=member_dict['name'],
            emailAddress=member_dict['email'],
        )


@dataclass(slots=True, frozen=True)
class Comment:
    body: str
    created_author: User
    created_timestring: str
    updated_author: User
    updated_timestring: str

    @classmethod
    def init_obj(cls, comment_obj: jira_res.Comment):
        return cls(
            body=utils.clean_string(comment_obj.body),
            created_author=User.init_obj(comment_obj.author),
            created_timestring=comment_obj.created,
            updated_author=User.init_obj(comment_obj.updateAuthor),
            updated_timestring=comment_obj.updated,
        )


@dataclass(slots=True, frozen=True)
class Worklog:
    id: str
    created_author: User
    created_timestring: str
    issueId: str
    comment: str
    started_timestring: str
    timeSpent: str
    timeSpentSeconds: int
    updated_author: User
    updated_timestring: str

    @classmethod
    def init_obj(cls, worklog_obj: jira_res.Worklog):
        return cls(
            id=worklog_obj.id,
            created_author=User.init_obj(worklog_obj.author),
            created_timestring=worklog_obj.created,
            issueId=worklog_obj.issueId,
            comment=worklog_obj.comment,
            started_timestring=worklog_obj.started,
            timeSpent=worklog_obj.timeSpent,
            timeSpentSeconds=worklog_obj.timeSpentSeconds,
            updated_author=User.init_obj(worklog_obj.updateAuthor),
            updated_timestring=worklog_obj.updated,
        )


@dataclass(slots=True, frozen=True)
class OptionValue:
    id: str
    value: str

    @classmethod
    def init_obj(cls, customfield_obj: jira_res.CustomFieldOption):
        return cls(
            id=customfield_obj.id,
            value=customfield_obj.value,
        )


@dataclass(slots=True, frozen=True)
class MultOptionValue:
    parent: OptionValue
    child: OptionValue

    @classmethod
    def init_obj(cls, customfield_obj: jira_res.CustomFieldOption):
        return cls(
            parent=OptionValue.init_obj(customfield_obj),
            child=OptionValue.init_obj(customfield_obj.child),
        )


@dataclass(slots=True, frozen=True)
class Component:
    id: str
    name: str

    @classmethod
    def init_obj(cls, components_obj: jira_res.Component):
        return cls(
            id=components_obj.id,
            name=components_obj.name,
        )


@dataclass(slots=True, frozen=True)
class Resolution:
    id: str
    name: str
    description: str

    @classmethod
    def init_obj(cls, resolution_obj: jira_res.Resolution):
        return cls(
            id=resolution_obj.id,
            name=resolution_obj.name,
            description=resolution_obj.description
        )
