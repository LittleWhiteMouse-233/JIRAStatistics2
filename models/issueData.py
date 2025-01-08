from abc import abstractmethod, ABC
import jira.resources as jira_res
import pandas as pd
from datetime import datetime
from typing import Callable, Any
from . import fieldStructure as fieldsS
from . import exceptions as exc
from . import utils


# 事务核心字段
class IssueLike:
    def __init__(self, issue_obj: jira_res.Issue):
        self.id = issue_obj.id
        self.key = issue_obj.key
        fields_obj = issue_obj.fields
        self.issueType = fieldsS.IssueType(fields_obj.issuetype)
        self.priority = fieldsS.Priority(fields_obj.priority)
        self.workflowStatus = fieldsS.WorkflowStatus(fields_obj.status)
        self.summary = utils.clean_string(fields_obj.summary)


# 任意事务类型
class Issue(IssueLike):
    def __init__(self, issue_obj: jira_res.Issue, ref_fields: fieldsS.FieldList):
        super().__init__(issue_obj)
        fields_obj = issue_obj.fields
        self.belongingProject = fieldsS.Project(fields_obj.project)
        description = fields_obj.description
        if description:
            self.description = utils.clean_string(description)
        else:
            self.description = ''
        # 用户类字段
        self.reporter = fieldsS.User(issue_obj.get_field('reporter'))
        self.creator = fieldsS.User(issue_obj.get_field('creator'))
        assignee = issue_obj.get_field('assignee')
        if assignee:
            self.assignee = fieldsS.User(assignee)
        else:
            self.assignee = None
        # 时间类字段
        self.created_timestring = issue_obj.get_field('created')
        self.updated_timestring = issue_obj.get_field('updated')
        # 标签类字段
        self.labels = fields_obj.labels
        self.components: list[fieldsS.Component] = []
        for component in issue_obj.get_field('components'):
            self.components.append(fieldsS.Component(component))
        # 评论列表
        self.comments = []
        for comment in fields_obj.comment.comments:
            self.comments.append(fieldsS.Comment(comment))
        # 工作日志
        self.worklogs = []
        for worklog in fields_obj.worklog.worklogs:
            self.worklogs.append(fieldsS.Worklog(worklog))
        # 子任务
        self.subtasks = []
        for subtask in issue_obj.get_field('subtasks'):
            self.subtasks.append(IssueLike(subtask))
        # 自定义字段
        self.base_platform: fieldsS.OptionValue = self.try_get_field(issue_obj, ref_fields.field_name2id('基础机芯&OS'),
                                                                     fieldsS.OptionValue)
        ## str or None
        self.other_platform: str = self.try_get_field(issue_obj, ref_fields.field_name2id('项目（其他）'), str)
        self.task_type: fieldsS.MultOptionValue = self.try_get_field(issue_obj, ref_fields.field_name2id('任务类型'),
                                                                     fieldsS.MultOptionValue)

    @classmethod
    def auto_adapt(cls, issue_obj: jira_res.Issue, ref_fields: fieldsS.FieldList):
        issue_type = issue_obj.fields.issuetype.name
        if issue_type == 'Epic':
            return Epic(issue_obj, ref_fields)
        elif issue_type == '任务':
            return Task(issue_obj, ref_fields)
        elif issue_type == '子任务':
            return Subtask(issue_obj, ref_fields)
        elif issue_type == '认证测试任务':
            return TestTask(issue_obj, ref_fields)
        elif issue_type == '认证管理任务':
            return ManageTask(issue_obj, ref_fields)
        else:
            return cls(issue_obj, ref_fields)

    @staticmethod
    def try_get_field(issue_obj: jira_res.Issue, field_name: str, class_type: type):
        try:
            field_obj = issue_obj.get_field(field_name)
        except AttributeError:
            return None
        if field_obj is None:
            return None
        else:
            return class_type(field_obj)

    def get_attribute(self, attr_name: str):
        if attr_name not in self.__dict__.keys():
            raise AttributeError("Issue has no attribute: %s." % attr_name)
        return self.__dict__[attr_name]

    @property
    def labels_string(self):
        return ', '.join(self.labels)

    @property
    def components_tuple(self):
        return tuple(map(lambda x: x.name, self.components))

    @property
    def components_string(self):
        return ', '.join(self.components_tuple)

    @property
    def colored_status(self):
        match str(self.workflowStatus.statusCategory):
            case '2':
                color = '<0xA6A6A6>'
            case '4':
                color = '<0x00B0F0>'
            case '3':
                color = '<0x00B050>'
            case _:
                color = ''
        return self.workflowStatus.name + color

    @property
    def total_workload(self):
        return sum(map(lambda x: x.timeSpentSeconds, self.worklogs))

    @property
    def platform_string(self):
        if self.base_platform is None or self.base_platform.value == 'Other':
            return self.other_platform
        else:
            return self.base_platform.value

    @property
    def info_string(self):
        return '[%s(%s)]%s' % (self.key, self.issueType.name, self.summary)

    @staticmethod
    def __generate_comment_string(comment: fieldsS.Comment, simple: bool = False, time_format: str = None):
        if simple:
            timestring = utils.parse_timestring(comment.created_timestring, time_format).strftime('%b %d, %a')
            return '>' * 3 + "%s, %s: \n%s" % (timestring, comment.created_author.displayName, comment.body)
        return "%s, %s(%s): \n%s" % (comment.created_timestring, comment.created_author.key,
                                     comment.created_author.emailAddress, comment.body)

    def generate_comments_series(self, begin: int = None, end: int = None):
        def stringification(x: fieldsS.Comment):
            return self.__generate_comment_string(x, simple=True)

        if self.comments:
            comment_list = list(map(stringification, self.comments[begin:end]))
        else:
            comment_list = ["### %s, Auto: \n### No comments." % datetime.now().strftime('%b %d, %a')]
        return pd.Series(comment_list)

    def get_comments_table(self, begin: int = None, end: int = None):
        comments_series = self.generate_comments_series(begin, end)
        summary = self.summary
        labels = "[%s(%s)]%s" % (self.key, self.issueType.name, self.labels_string)
        status = self.colored_status
        comments_table = utils.concat_single_value(comments_series,
                                                   left=[self.belongingProject.name, summary, labels, status],
                                                   repeat=True,
                                                   columns=['project', 'summary', 'labels', 'status', 'comments'])
        return comments_table


# Epic 型事务
class Epic(Issue):
    def __init__(self, issue_obj: jira_res.Issue, ref_fields: fieldsS.FieldList):
        super().__init__(issue_obj, ref_fields)
        # Epic 专属字段
        self.epic_name = issue_obj.get_field(ref_fields.field_name2id('Epic Name'))
        # 自定义字段
        ## 级联列表
        self.certification: fieldsS.MultOptionValue = self.try_get_field(issue_obj, ref_fields.field_name2id('认证项'),
                                                                         fieldsS.MultOptionValue)

    @property
    def certification_string(self):
        if self.certification is None:
            return None
        return '%s-%s' % (self.certification.parent.value, self.certification.child.value)

    def get_comments_table(self, begin: int = None, end: int = None):
        comments_series = self.generate_comments_series(begin, end)
        summary = self.platform_string if self.platform_string is not None else self.summary
        labels = "[%s(%s)]%s" % (self.key, self.issueType.name,
                                 self.certification_string if self.certification else self.labels_string)
        status = self.colored_status
        comments_table = utils.concat_single_value(comments_series,
                                                   [self.belongingProject.name, summary, labels, status],
                                                   [],
                                                   repeat=True,
                                                   columns=['project', 'summary', 'labels', 'status', 'comments'])
        return comments_table


# 类任务型事务
class TaskLike(ABC, Issue):
    def __init__(self, issue_obj: jira_res.Issue, ref_fields: fieldsS.FieldList):
        super().__init__(issue_obj, ref_fields)
        self.epic_link = None
        self.parent = None

    @abstractmethod
    def generate_coordinate(self, epic: Issue, *, task: Issue = None):
        pass

    @property
    def parent_key(self):
        if issubclass(self.__class__, Task):
            self.epic_link: str
            return self.epic_link
        if issubclass(self.__class__, Subtask):
            self.parent: IssueLike
            return self.parent.key


# 任务型事务
class Task(TaskLike):
    def __init__(self, issue_obj: jira_res.Issue, ref_fields: fieldsS.FieldList):
        super().__init__(issue_obj, ref_fields)
        # 任务专属字段
        self.epic_link = issue_obj.get_field(ref_fields.field_name2id('Epic Link'))

    def verify(self, epic: Epic, task: TaskLike):
        assert epic.key == self.epic_link
        if task is not None:
            assert self.key == task.key

    def generate_coordinate(self, epic: Epic, task: TaskLike = None):
        self.verify(epic, task)
        r1 = epic.epic_name
        r2 = None
        c1 = '无标定任务'
        c2 = None
        return r1, r2, c1, c2


# 认证测试任务
class TestTask(Task):
    def generate_coordinate(self, epic: Epic, task: Task = None):
        self.verify(epic, task)
        if epic.certification is None:
            raise exc.InvalidFieldError('certification')
        r1 = epic.certification.child.value
        if not self.components:
            raise exc.InvalidFieldError('components')
        r2 = self.components_tuple
        if self.task_type is None:
            raise exc.InvalidFieldError('task_type')
        c1 = self.task_type.parent.value
        c2 = self.task_type.child.value
        return r1, r2, c1, c2


# 认证管理任务
class ManageTask(Task):
    def generate_coordinate(self, epic: Epic, task: Task = None):
        self.verify(epic, task)
        if epic.certification is None:
            raise exc.InvalidFieldError('certification')
        r1 = epic.certification.child.value
        if not self.components:
            raise exc.InvalidFieldError('components')
        r2 = self.components_tuple
        c1 = "认证管理"
        c2 = None
        return r1, r2, c1, c2


# 子任务型事务
class Subtask(TaskLike):
    def __init__(self, issue_obj: jira_res.Issue, ref_fields: fieldsS.FieldList):
        super().__init__(issue_obj, ref_fields)
        # 子任务专属字段
        self.parent = IssueLike(issue_obj.get_field('parent'))

    def __gen_coord_as_test(self, epic: Epic, task: Task):
        if epic.certification is None:
            raise exc.InvalidFieldError('certification')
        r1 = epic.certification.child.value
        if not self.components:
            raise exc.InvalidFieldError('components')
        r2 = self.components_tuple
        if self.task_type is not None:
            c1 = self.task_type.parent.value
            c2 = self.task_type.child.value
        elif task is not None and task.task_type is not None:
            c1 = task.task_type.parent.value
            c2 = task.task_type.child.value
        else:
            raise exc.InvalidFieldError('task_type')
        return r1, r2, c1, c2

    def __gen_coord_as_manage(self, epic: Epic):
        if epic.certification is None:
            raise exc.InvalidFieldError('certification')
        r1 = epic.certification.child.value
        if not self.components:
            raise exc.InvalidFieldError('components')
        r2 = self.components_tuple
        c1 = '认证管理'
        c2 = None
        return r1, r2, c1, c2

    @staticmethod
    def __gen_coord_as_public(epic: Epic):
        r1 = epic.epic_name
        r2 = None
        c1 = '无标定子任务'
        c2 = None
        return r1, r2, c1, c2

    def verify(self, epic: Epic, task: Task):
        if task is None:
            assert epic.key == self.parent.key
        else:
            assert task.key == self.parent.key
            assert epic.key == task.epic_link

    def generate_coordinate(self, epic: Epic, task: Task = None):
        self.verify(epic, task)
        if type(task) is TestTask:
            return self.__gen_coord_as_test(epic, task)
        if type(task) is ManageTask:
            return self.__gen_coord_as_manage(epic)
        return self.__gen_coord_as_public(epic)


class IssueList(list[Issue]):
    def __init__(self, field_obj_list: list[dict[str, Any]] = None):
        super().__init__()
        self.__ref_fields = None
        if field_obj_list is not None:
            self.__ref_fields = fieldsS.FieldList(field_obj_list)

    def import_issues(self, issue_obj_list: list[jira_res.Issue]):
        if self.__ref_fields is None:
            raise ValueError("This instance does not have a FieldList for reference, can not import issues.")
        for issue_obj in issue_obj_list:
            issue = Issue.auto_adapt(issue_obj, self.__ref_fields)
            print("Import issue: [%s(%s)]%s." % (issue.key, issue.issueType.name, issue.summary))
            self.append(issue)
        print("Import completed! (Total=%d)\n" % len(issue_obj_list))

    def get_comments_status(self, begin: int = None, end: int = None, filter_func: Callable[[Issue], bool] = None):
        concat_list = []
        for issue in self:
            if filter_func is not None and not filter_func(issue):
                continue
            concat_list.append(issue.get_comments_table(begin, end))
        comments_table = pd.concat(concat_list)
        comments_table.sort_values(by=list(comments_table.columns[[0, 1, 2]]), inplace=True)
        comments_table.reset_index(drop=True, inplace=True)
        return comments_table

    def __listing_attribute(self, func: Callable[[Issue], str], assert_unique=False):
        attr_list = list(map(func, self))
        if assert_unique:
            assert len(attr_list) == len(set(attr_list))
            assert None not in attr_list
        return attr_list

    @property
    def key_list(self):
        return self.__listing_attribute(lambda x: x.key, assert_unique=True)

    @property
    def id_list(self):
        return self.__listing_attribute(lambda x: x.id, assert_unique=True)

    def self_search_by(self, key_or_id: str, return_index=False):
        if key_or_id in self.key_list:
            index = self.key_list.index(key_or_id)
        elif key_or_id in self.id_list:
            index = self.id_list.index(key_or_id)
        else:
            return None
        if return_index:
            return index
        else:
            return self[index]

    def has(self, key_or_id: str):
        if self.self_search_by(key_or_id) is not None:
            return True
        else:
            return False


def get_attribute_by_queue(attr_name: str, issue_queue: list[Issue]):
    attr = None
    for issue in issue_queue:
        if issue is None:
            continue
        attr = issue.get_attribute(attr_name)
        if attr is not None:
            break
    return attr
