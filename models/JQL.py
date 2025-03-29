from enum import Enum
from dataclasses import dataclass


@dataclass(slots=True)
class JQLFilter:
    description: str
    content: str


def or_(jql_f1: JQLFilter, jql_f2: JQLFilter):
    return "(%s) OR (%s)" % (jql_f1.content, jql_f2.content)


def and_(jql_f1: JQLFilter, jql_f2: JQLFilter):
    return "(%s) AND (%s)" % (jql_f1.content, jql_f2.content)


class BaseFilter(JQLFilter, Enum):
    UNFINISHED_EPIC = (
        r"非谷歌未完成 Epic 按项目-优先级-概要排序",
        r"project in (DT, PAR, VOD) "
        r"AND issuetype = Epic "
        r"AND resolution in (Unresolved, IP-后处理阶段) "
        r"ORDER BY project, level, summary"
    )
    RECENTLY_RESOLVED_EPIC = (
        r"非谷歌最近解决的Epic",
        r"project in (DTCER, PARCER, VODCER) AND issuetype = Epic AND resolved >= -2w"
    )
    UNRESOLVED_EPIC = (
        r"非谷歌未解决的Epic",
        r"project in (DTCER, PARCER, VODCER) AND issuetype = Epic AND resolution = Unresolved"
    )
    ALL_TASK_LIKE = (
        r"非谷歌所有任务和子任务",
        r"project in (DTCER, PARCER, VODCER) AND issuetype in (认证测试任务, 认证管理任务, Sub-task)"
    )

    ALL_TASK_LIKE_GOOGLE = (
        r"谷歌所有任务和子任务",
        r"project in (GTVS) AND issuetype in (认证测试任务, 认证管理任务, Sub-task)"
    )


class ConcatFilter(JQLFilter, Enum):
    EPIC_COMMENT = (
        r"Epic 评论",
        or_(BaseFilter.RECENTLY_RESOLVED_EPIC, BaseFilter.UNRESOLVED_EPIC)
    )
    ALL_TASK_LIKE = (
        r"所有任务和子任务",
        or_(BaseFilter.ALL_TASK_LIKE, BaseFilter.ALL_TASK_LIKE_GOOGLE)
    )
