import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
import pandas as pd
import numpy as np
from enum import Enum
from . import issueData as issueD
from . import fieldStructure as fieldS
from . import exceptions as exc
from . import utils
from .accessAgent import JIRAOperator
from .workbookProcess import WorksheetProcessor


class Workload:
    class __EmptyWorklog:
        def __init__(self, issue: issueD.Issue):
            self.issueId = issue.id
            self.created_author = issue.assignee
            if self.created_author is None:
                self.created_author = issue.creator
            self.comment = "# EmptyWorklog"
            self.timeSpentSeconds = 0

    def __init__(self, worklog: fieldS.Worklog | __EmptyWorklog, jira_op: JIRAOperator, rate: float):
        issue = jira_op.find_issue_by(worklog.issueId)
        task, epic = jira_op.find_parents(issue)
        self.__issue: issueD.Issue = issue
        self.__task: issueD.Task = task
        self.__epic: issueD.Epic = epic
        self.__worklog = worklog
        self.__rate = rate
        # 额外信息
        self.__extend_of_issue(issue)
        self.__extend_of_task(task)
        self.__extend_of_epic(epic)

    @classmethod
    def empty_workload(cls, issue: issueD.Issue, jira_op: JIRAOperator, rate: float):
        return cls(cls.__EmptyWorklog(issue), jira_op, rate)

    def __extend_of_issue(self, issue: issueD.Issue):
        self.project = issue.belongingProject
        self.task_type = issue.task_type

    def __extend_of_task(self, task: issueD.Task):
        if task is not None:
            self.platform = task.platform_string
            if self.task_type is None:
                self.task_type = task.task_type
        else:
            self.platform = None

    def __extend_of_epic(self, epic: issueD.Epic):
        self.certification = epic.certification_string
        if self.platform is None:
            self.platform = epic.platform_string

    @property
    def belong_issue_key(self):
        return self.__issue.key

    def is_empty(self):
        return type(self.__worklog) is self.__EmptyWorklog

    def get_worklog_info(self):
        return pd.Series([
            self.__issue.info_string,
            self.__task.info_string if self.__task else '# NoTask',
            self.__epic.info_string,
            self.project.name,
            self.__worklog.created_author.displayName,
            self.__worklog.comment,
            self.__worklog.timeSpentSeconds,
            self.__rate,
        ])


class ReferenceMap:
    def __init__(self, worksheet: Worksheet, origin_point_rc: tuple[int, int]):
        # 输入值为坐标重叠域右下交点的自然行列坐标，对应值域左上起点的零点行列坐标
        self.__op_r, self.__op_c = origin_point_rc
        self.__name = worksheet.title
        filled_sheet = WorksheetProcessor(worksheet).unmerge_cells_and_fill(fill=True)
        self.__map = pd.DataFrame(filled_sheet.values)
        self.__map.dropna(axis=0, how='all', inplace=True)
        self.__map.dropna(axis=1, how='all', inplace=True)

    @property
    def __level_x(self):
        return self.__op_r

    @property
    def __level_y(self):
        return self.__op_c

    @property
    def __value_map(self):
        return self.__map.iloc[self.__op_r:, self.__op_c:].reset_index(drop=True).T.reset_index(drop=True).T

    @property
    def value_shape(self):
        return self.__map.shape[0] - self.__op_r, self.__map.shape[1] - self.__op_c

    @property
    def worksheet_name(self):
        return self.__name

    def __locate_coord_value(self, value: str, level: int, axis: int):
        if axis == 0:
            if level > self.__level_x:
                raise ValueError("The level(%d) is over than level-num of x axis(%d)." % (level, self.__level_x))
            return self.__map.iloc[level - 1, self.__level_y:].eq(value)
        elif axis == 1:
            if level > self.__level_y:
                raise ValueError("The level(%d) is over than level-num of y axis(%d)." % (level, self.__level_y))
            return self.__map.iloc[self.__level_x:, level - 1].eq(value)
        else:
            raise ValueError("The axis should be 0 or 1, but get %d." % axis)

    def __locate_multilayer_coord(self, coord_list: list[str], axis: int, auto_adapt=False):
        if auto_adapt:
            coord_list = self.__auto_adapt_coord_list(coord_list, axis)
        level_xy = (self.__level_x, self.__level_y)
        if len(coord_list) != level_xy[axis]:
            raise ValueError("The length of coord_list: %d is different from the level-num of axis(%d): %d."
                             % (len(coord_list), axis, level_xy[axis]))
        bool_index = pd.Series([True] * self.value_shape[int(not bool(axis))])
        for i, coord in enumerate(coord_list):
            if coord is None:
                continue
            bool_index &= self.__locate_coord_value(coord, level=i + 1, axis=axis)
            if not bool_index.any():
                break
        return bool_index

    def __auto_adapt_coord_list(self, coord_list: list[str], axis: int):
        i = len(coord_list)
        for i in reversed(range(len(coord_list))):
            if coord_list[i] is not None:
                break
        ata_coord_list = coord_list[:i + 1]
        level_length = (self.__level_x, self.__level_y)[axis]
        return [None] * (level_length - len(ata_coord_list)) + ata_coord_list

    def locate_coord_cell(self, row_coordinates: list[str], col_coordinates: list[str]):
        row_index = self.__locate_multilayer_coord(row_coordinates, axis=1, auto_adapt=True)
        col_index = self.__locate_multilayer_coord(col_coordinates, axis=0, auto_adapt=True)
        located = self.__value_map.loc[row_index, col_index]
        coord = (*row_coordinates, *col_coordinates)
        if located.empty:
            raise exc.NoMatchingError(coord)
        if located.size != 1:
            raise exc.ManyMatchingError(coord)
        if pd.isna(located.iloc[0, 0]):
            raise exc.MatchingNAError(coord)
        return located

    @staticmethod
    def cell2value(located: pd.DataFrame):
        return located.iloc[0, 0]

    @staticmethod
    def cell2index(located: pd.DataFrame):
        return int(located.index[0]), int(located.columns[0])

    def value_map2table(self, value_map: np.ndarray):
        value_table = pd.DataFrame(value_map)
        table_x = self.__map.iloc[:self.__level_x, self.__level_y:].T.reset_index(drop=True).T
        value_x = pd.concat([table_x, value_table]).reset_index(drop=True)
        table_y = self.__map.iloc[:, :self.__level_y]
        value_xy = pd.concat([table_y, value_x], axis=1)
        return value_xy


class Cell:
    def __init__(self, r1: str, r2: str, c1: str, c2: str, ref_map: ReferenceMap):
        self.__workloads: list[Workload] = []
        self.__r1 = r1
        self.__r2 = r2
        self.__c1 = c1
        self.__c2 = c2
        self.__ref_map = ref_map
        located_cell = ref_map.locate_coord_cell([self.__r1, self.__r2], [self.__c1, self.__c2])
        self.__std_time = ReferenceMap.cell2value(located_cell)
        self.__value_index = ReferenceMap.cell2index(located_cell)

    def add_workload(self, issue: issueD.Issue, jira_op: JIRAOperator, rate: float):
        if issue.worklogs:
            for worklog in issue.worklogs:
                assert worklog.issueId == issue.id
                self.__workloads.append(Workload(worklog, jira_op, rate))
        else:
            self.__workloads.append(Workload.empty_workload(issue, jira_op, rate))

    @property
    def coord_string(self):
        return '([%s]-[%s], [%s]-[%s])' % self.coord_tuple

    @property
    def coord_tuple(self):
        return self.__r1, self.__r2, self.__c1, self.__c2

    @property
    def num_worklog(self):
        return len(self.__workloads)

    @property
    def num_issues(self):
        return len(set(map(lambda x: x.belong_issue_key, self.__workloads)))

    @property
    def value_index(self):
        return self.__value_index

    def refer_from(self, ref_map: ReferenceMap):
        if ref_map is self.__ref_map or ref_map.worksheet_name == self.__ref_map.worksheet_name:
            return True
        else:
            return False

    def get_worklog_table(self):
        concat_list = []
        for workload in self.__workloads:
            concat_list.append(workload.get_worklog_info())
        table = pd.concat(concat_list, axis=1).T
        table = utils.concat_single_value(table,
                                          left=[self.coord_string],
                                          repeat=True,
                                          columns=['coordinate', 'issue', 'task', 'epic', 'project',
                                                   'creator', 'comment', 'time(sec)', 'rate'])
        return table


class Matrix:
    class LoadResult(Enum):
        SKIP = 1, 'Skip'
        SUCCESS = 0, 'Success'
        WRONG = -1, 'Wrong'

        def __init__(self, result_id: int, result_name: str):
            self.res_id = result_id
            self.res_name = result_name

    class __IssueRecord:
        def __init__(self, issue: issueD.Issue | issueD.TaskLike):
            self.__issue = issue
            self.__result = None
            self.__detail = None
            self.ref_class = type(issue)

        @property
        def issue(self):
            return self.__issue

        @property
        def result(self):
            return self.__result

        @property
        def detail(self):
            return self.__detail

        def skip(self, detail: str = None):
            self.__result = Matrix.LoadResult.SKIP
            self.__detail = detail

        def success(self, detail: str = None):
            self.__result = Matrix.LoadResult.SUCCESS
            self.__detail = detail

        def wrong(self, detail: str = None):
            self.__result = Matrix.LoadResult.WRONG
            self.__detail = detail

    def __init__(self, issues: issueD.IssueList, jira_op: JIRAOperator, ref_filename: str):
        jira_op.add_cache(issues)
        self.__jira_op = jira_op
        ref_xlsx = openpyxl.load_workbook(ref_filename)
        self.__ref_test = ReferenceMap(ref_xlsx.worksheets[0], (2, 3))
        self.__ref_manage = ReferenceMap(ref_xlsx.worksheets[1], (2, 3))
        self.__records: list[Matrix.__IssueRecord] = []
        self.__cells: list[Cell] = []
        for issue in issues:
            self.__records.append(self.__IssueRecord(issue))
        for i_r in self.__records:
            issue = i_r.issue
            # 非类任务型
            if not issubclass(type(issue), issueD.TaskLike):
                i_r.skip("Is not subclass of TaskLike: %s." % issue.issueType.name)
                continue
            index = issues.self_search_by(issue.parent_key, return_index=True)
            # 排除存在子任务的父任务
            if index is not None:
                i_r_p = self.__records[index]
                i_r_p.skip("This issue is parent of: %s." % i_r.issue.key)
        self.load_workload_into_cell()

    @staticmethod
    def __coordinate_generator(coord: tuple):
        def ranging(rc: str | tuple[str] | list[str]):
            if type(rc) is tuple or type(rc) is list:
                return rc
            return (rc,)

        r1, r2, c1, c2 = coord
        for i1 in ranging(r1):
            for i2 in ranging(r2):
                for j1 in ranging(c1):
                    for j2 in ranging(c2):
                        i1: str
                        i2: str
                        j1: str
                        j2: str
                        yield i1, i2, j1, j2

    def __analyse_coordinate(self, issue_record: __IssueRecord):
        task_like = issue_record.issue
        # 查找父事务
        task, epic = self.__jira_op.find_parents(task_like)
        if type(task_like) is issueD.Subtask and issubclass(type(task), issueD.TaskLike):
            issue_record.ref_class = type(task)
        # 生成坐标（集合）
        coord = task_like.generate_coordinate(epic, task=task)
        # 拆分坐标集合（例如“模块”字段），构建坐标生成器
        coord_generator = self.__coordinate_generator(coord)
        return coord_generator

    def load_workload_into_cell(self):
        print("Loading workload into cell ...")
        for i_r in self.__records:
            if i_r.result is self.LoadResult.SKIP:
                continue
            try:
                coord_generator = self.__analyse_coordinate(i_r)
            # 找不到父级事务
            except exc.GetIssueFailedError as e:
                i_r.wrong("%s(TaskLike.parent_key=%s)." % (e, i_r.issue.parent_key))
                continue
            # 坐标获取异常
            except exc.CoordinateError as e:
                i_r.wrong(str(e))
                continue
            cell_list = []
            try:
                for coord in coord_generator:
                    cell = self.__find_cell_or_create(coord, i_r.ref_class)
                    cell_list.append(cell)
            # 匹配坐标（集合）失败
            except exc.MisMatchingError as e:
                i_r.wrong(str(e))
                continue
            for cell in cell_list:
                cell.add_workload(i_r.issue, self.__jira_op, 1 / len(cell_list))
            i_r.success("Coordinate(s): " + ' & '.join([cell.coord_string for cell in cell_list]))
        assert len(self.__records) == sum(self.__num_of(res) for res in self.LoadResult)
        print("Loading issue completed.\n")

    def __find_cell_or_create(self, ref_coord: tuple[str, str, str, str], ref_class: type):
        flag = True
        for cell in self.__cells:
            coord = cell.coord_tuple
            for i in range(len(coord)):
                if coord[i] != ref_coord[i]:
                    flag = False
                    break
            if flag:
                return cell
            flag = True
        # 已有的 Cell 没有坐标能对应上，新建 Cell
        if ref_class is issueD.TestTask:
            ref_map = self.__ref_test
        elif ref_class is issueD.ManageTask:
            ref_map = self.__ref_manage
        else:
            ref_map = None
        new_cell = Cell(*ref_coord, ref_map=ref_map)
        self.__cells.append(new_cell)
        return new_cell

    def export_worklog_table(self):
        concat_list = []
        for cell in self.__cells:
            concat_list.append(cell.get_worklog_table())
        table = pd.concat(concat_list).reset_index(drop=True)
        table.sort_values(by=list(table.columns[[0, 1, 4]]), inplace=True)
        return table

    def __num_of(self, res: LoadResult):
        return sum(map(lambda x: x.result is res, self.__records))

    def running_report(self, show_detail=True):
        print("Run report: ")
        if show_detail:
            for res in self.LoadResult:
                print("\t%s(%s): " % (res.res_name, self.__num_of(res)))
                for rec in self.__records:
                    if rec.result is res:
                        print(utils.specific_length_string(rec.issue.info_string), rec.detail)
        print('Total: %d' % len(self.__records), end=', ')
        print(', '.join(map(lambda x: '%s: %d' % (x.res_name, self.__num_of(x)), self.LoadResult)))

    def __build_count_matrix(self, ref_map: ReferenceMap):
        array = np.zeros(ref_map.value_shape)
        for cell in self.__cells:
            if not cell.refer_from(ref_map):
                continue
            row_index, col_index = cell.value_index
            array[row_index][col_index] = cell.num_issues
        return array

    def export_count_matrix(self, ref_map: ReferenceMap):
        value_map = self.__build_count_matrix(ref_map)
        value_table = ref_map.value_map2table(value_map)
        return value_table

    def export_workload_matrix(self):
        pass
