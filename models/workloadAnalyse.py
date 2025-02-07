import pandas as pd
import numpy as np
import warnings
from enum import Enum
from typing import Callable
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet
# from copy import deepcopy
from .accessAgent import JIRAOperator
from .workbookProcess import WorksheetProcessor, HeatmapRenderer, RCActivator
from . import issueData as issueD
from . import fieldStructure as fieldS
from . import exceptions as exc
from . import utils


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

    @property
    def person_hour(self):
        return self.__rate * self.__worklog.timeSpentSeconds / 3600

    @property
    def person_day(self):
        return self.person_hour / 8

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
            self.__worklog.timeSpentSeconds / 3600,
            self.__rate,
        ])


class ReferenceMap:
    def __init__(self, worksheet: Worksheet, origin_point_rc: tuple[int, int]):
        # 输入值为坐标重叠域右下交点的自然行列坐标，对应值域左上起点的零点行列坐标
        self.__origin_ws = worksheet
        # 原点行坐标，原点列坐标
        self.__op_r, self.__op_c = origin_point_rc
        ws = WorksheetProcessor.copy_into(worksheet, Workbook().active)
        unmerged_sheet = WorksheetProcessor.unmerge_cells_and_fill(ws, fill=True)
        # unmerged_sheet = WorksheetProcessor(deepcopy(worksheet)).unmerge_cells_and_fill(fill=True)
        table = pd.DataFrame(unmerged_sheet.values)
        table.dropna(axis=0, how='all', inplace=True)
        table.dropna(axis=1, how='all', inplace=True)
        self.__value_map = self.__reset_rc(table.iloc[self.__op_r:, self.__op_c:])
        self.__axis_x = self.__reset_rc(table.iloc[:self.__op_r, self.__op_c:])
        self.__axis_y = self.__reset_rc(table.iloc[self.__op_r:, :self.__op_c])

    @staticmethod
    def __reset_rc(df: pd.DataFrame, row=True, col=True, row_drop=True, col_drop=True):
        if row:
            df = df.reset_index(drop=row_drop)
        if col:
            df = df.T.reset_index(drop=col_drop).T
        return df

    # 横轴层数
    @property
    def __ln_x(self):
        return self.__op_r

    # 纵轴层数
    @property
    def __ln_y(self):
        return self.__op_c

    @property
    def value_shape(self):
        return self.__value_map.shape

    @property
    def worksheet_name(self):
        return self.__origin_ws.title

    def __check_level_x(self, level_x: int):
        if not 1 <= level_x <= self.__ln_x:
            raise ValueError("The level_x(%d) is out of level-num range(1, %d)." % (level_x, self.__ln_x))

    def __check_level_y(self, level_y: int):
        if not 1 <= level_y <= self.__ln_y:
            raise ValueError("The level_y(%d) is out of level-num range(1, %d)." % (level_y, self.__ln_y))

    def __locate_coord_value(self, value: str, level: int, axis: int):
        # level: belong to N*
        if axis == 0:
            self.__check_level_x(level)
            return self.__axis_x.iloc[level - 1, :].eq(value)
        elif axis == 1:
            self.__check_level_y(level)
            return self.__axis_y.iloc[:, level - 1].eq(value)
        else:
            raise ValueError("The axis should be 0 or 1, but get %d." % axis)

    def __locate_multilayer_coord(self, coord_list: list[str], axis: int, auto_adapt=False):
        if auto_adapt:
            coord_list = self.__auto_adapt_coord_list(coord_list, axis)
        xy_ln = (self.__ln_x, self.__ln_y)
        if len(coord_list) != xy_ln[axis]:
            raise ValueError("The length of coord_list: %d is different from the level-num of axis(%d): %d."
                             % (len(coord_list), axis, xy_ln[axis]))
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
        level_length = (self.__ln_x, self.__ln_y)[axis]
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
        located_value = pd.to_numeric(located.iloc[0, 0], errors='coerce')
        if pd.isna(located_value):
            raise exc.MatchingNAError(coord)
        return located

    @staticmethod
    def cell2value(located: pd.DataFrame):
        return pd.to_numeric(located.iloc[0, 0])

    @staticmethod
    def cell2index(located: pd.DataFrame):
        return int(located.index[0]), int(located.columns[0])

    def value_array2sheet(self, value_array: np.ndarray, heatmap=True):
        # 数组的行列数应该与参考表的数据矩阵的行列数一致
        if value_array.shape != self.value_shape:
            raise ValueError("The shape of value_array(%s) is wrong, it shall be the same as ref_map: %s)."
                             % value_array.shape, self.value_shape)
        worksheet = WorksheetProcessor.copy_into(self.__origin_ws, Workbook().active)
        # 渲染器
        renderer = HeatmapRenderer(worksheet, value_array.max(), value_array.min()) if heatmap else None
        for i in range(value_array.shape[0]):
            for j in range(value_array.shape[1]):
                value = value_array[i][j]
                ws_i = self.__ln_x + i + 1
                ws_j = self.__ln_y + j + 1
                # 热力图
                if heatmap:
                    renderer.colorful_value(ws_i, ws_j, value)
                else:
                    worksheet.cell(ws_i, ws_j).value = value
        return worksheet

    def value_array2downmix_sheet(self, value_array: np.ndarray, level_x: int = None, level_y: int = None,
                                  heatmap=True):
        # level_x, level_y: belong to N*
        if level_x is None and level_y is None:
            warnings.warn("If you want a sheet with the same shape as the ref_map, use value_array2sheet().",
                          RuntimeWarning)
            return self.value_array2sheet(value_array, heatmap=heatmap)
        # 数组的行列数应该与参考表的数据矩阵的行列数一致
        if value_array.shape != self.value_shape:
            raise ValueError("The shape of value_array(%s) is wrong, it shall be the same as ref_map: %s)."
                             % value_array.shape, self.value_shape)

        def downmix(array: np.ndarray, axis_series: pd.Series, axis: int):
            assert axis == 0 or axis == 1
            assert axis_series.size == array.shape[-(axis + 1)]
            concat_list = []
            dm_axis = [axis_series[0]]
            start = 0
            flag = False
            for end in range(1, axis_series.size + 1):
                if end == axis_series.size:
                    flag = True
                else:
                    coord = axis_series[end]
                    if coord != dm_axis[-1]:
                        dm_axis.append(coord)
                        flag = True
                if flag:
                    assert end > start
                    if axis == 0:
                        concat_list.append(array[:][start:end].sum(axis=1))
                    else:
                        concat_list.append(array[start:end][:].sum(axis=0))
                    start = end
                    flag = False
            assert len(concat_list) == len(dm_axis)
            if axis == 0:
                return np.vstack(concat_list).T, pd.DataFrame(dm_axis).T
            else:
                return np.vstack(concat_list), pd.DataFrame(dm_axis)

        if level_x is not None:
            self.__check_level_x(level_x)
            downmix_x, axis_x = downmix(value_array, self.__axis_x.iloc[level_x - 1, :], axis=0)
        else:
            downmix_x = value_array
            axis_x = self.__axis_x
        if level_y is not None:
            self.__check_level_y(level_y)
            downmix_xy, axis_y = downmix(downmix_x, self.__axis_y.iloc[:, level_y - 1], axis=1)
        else:
            downmix_xy = downmix_x
            axis_y = self.__axis_y
        op_r, op_c = axis_x.shape[0], axis_y.shape[1]
        worksheet = Workbook().active
        # 填充横坐标域
        if level_x is not None:
            for i in range(axis_x.shape[0]):
                for j in range(axis_x.shape[1]):
                    worksheet.cell(i + 1, op_c + j + 1).value = axis_x.iloc[i, j]
        else:
            # 复制原表的横坐标域
            WorksheetProcessor.copy_part_into(self.__origin_ws, worksheet,
                                              RCActivator.scope_int2str(1,
                                                                        self.__op_c + 1,
                                                                        self.__op_r,
                                                                        self.__op_c + self.__axis_x.shape[1]),
                                              RCActivator.point_int2str(1, op_c + 1))
        # 填充纵坐标域
        if level_y is not None:
            for i in range(axis_y.shape[0]):
                for j in range(axis_y.shape[1]):
                    worksheet.cell(op_r + i + 1, j + 1).value = axis_y.iloc[i, j]
        else:
            # 复制原表的纵坐标域
            WorksheetProcessor.copy_part_into(self.__origin_ws, worksheet,
                                              RCActivator.scope_int2str(self.__op_r + 1,
                                                                        1,
                                                                        self.__op_r + self.__axis_y.shape[0],
                                                                        self.__op_c),
                                              RCActivator.point_int2str(op_r + 1, 1))
        # 补充表头域
        if level_x is None and level_y is not None:
            # 复制压缩列对应的表头域
            WorksheetProcessor.copy_part_into(self.__origin_ws, worksheet,
                                              RCActivator.scope_int2str(1,
                                                                        level_y,
                                                                        self.__op_r,
                                                                        level_y),
                                              RCActivator.point_int2str(1, 1))
        # 填充值域
        renderer = HeatmapRenderer(worksheet, value_array.max(), value_array.min()) if heatmap else None
        for i in range(downmix_xy.shape[0]):
            for j in range(downmix_xy.shape[1]):
                ws_i, ws_j = op_r + i + 1, op_c + j + 1
                value = downmix_xy[i][j]
                if heatmap:
                    renderer.colorful_value(ws_i, ws_j, value)
                else:
                    worksheet.cell(ws_i, ws_j).value = value
        return worksheet


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
    def coord_index(self):
        return self.__value_index

    @property
    def num_worklog(self):
        return len(self.__workloads)

    @property
    def num_issues(self):
        return len(set(map(lambda x: x.belong_issue_key, self.__workloads)))

    def cumulative_workload(self, unit='day'):
        if unit == 'day':
            return sum(map(lambda x: x.person_day, self.__workloads))
        elif unit == 'hour':
            return sum(map(lambda x: x.person_hour, self.__workloads))
        else:
            raise ValueError("Unknown unit: %s." % unit)

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
                                                   'creator', 'comment', 'time(hour)', 'rate'])
        return table


class Matrix:
    class LoadResult(Enum):
        SKIP = 1, 'Skip'
        SUCCESS = 0, 'Success'
        WRONG = -1, 'Wrong'

        def __init__(self, result_id: int, result_name: str):
            self.res_id = result_id
            self.res_name = result_name

    class __MetaData:
        def __init__(self, issue: issueD.Issue | issueD.TaskLike):
            self.__issue = issue
            self.__load_result = None
            self.___load_detail = None
            self.ref_class = type(issue)

        @property
        def issue(self):
            return self.__issue

        @property
        def load_result(self):
            return self.__load_result

        @property
        def load_detail(self):
            return self.___load_detail

        def skip(self, detail: str = None):
            self.__load_result = Matrix.LoadResult.SKIP
            self.___load_detail = detail

        def success(self, detail: str = None):
            self.__load_result = Matrix.LoadResult.SUCCESS
            self.___load_detail = detail

        def wrong(self, detail: str = None):
            self.__load_result = Matrix.LoadResult.WRONG
            self.___load_detail = detail

    def __init__(self, issues: issueD.IssueList, jira_op: JIRAOperator, ref_filename: str):
        jira_op.add_cache(issues)
        self.__jira_op = jira_op
        ref_xlsx = load_workbook(ref_filename)
        self.__ref_test = ReferenceMap(ref_xlsx.worksheets[0], (2, 3))
        self.__ref_manage = ReferenceMap(ref_xlsx.worksheets[1], (2, 3))
        self.__meta_datas: list[Matrix.__MetaData] = []
        self.__cells: list[Cell] = []
        for issue in issues:
            self.__meta_datas.append(self.__MetaData(issue))
        for metadata in self.__meta_datas:
            issue = metadata.issue
            # 非类任务型
            if not issubclass(type(issue), issueD.TaskLike):
                metadata.skip("Is not subclass of TaskLike: %s." % issue.issueType.name)
                continue
            index = issues.self_search_by(issue.parent_key, return_index=True)
            # 排除存在子任务的父任务
            if index is not None:
                md_p = self.__meta_datas[index]
                md_p.skip("This issue is parent of: %s." % metadata.issue.key)
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

    def __analyse_coordinate(self, metadata: __MetaData):
        task_like = metadata.issue
        # 查找父事务
        task, epic = self.__jira_op.find_parents(task_like)
        if type(task_like) is issueD.Subtask and issubclass(type(task), issueD.TaskLike):
            metadata.ref_class = type(task)
        # 生成坐标（集合）
        coord = task_like.generate_coordinate(epic, task=task)
        # 拆分坐标集合（例如“模块”字段），构建坐标生成器
        coord_generator = self.__coordinate_generator(coord)
        return coord_generator

    def load_workload_into_cell(self):
        print("Loading workload into cell ...")
        for metadata in self.__meta_datas:
            if metadata.load_result is self.LoadResult.SKIP:
                continue
            try:
                coord_generator = self.__analyse_coordinate(metadata)
            # 找不到父级事务
            except exc.GetIssueFailedError as e:
                metadata.wrong("%s(TaskLike.parent_key=%s)." % (e, metadata.issue.parent_key))
                continue
            # 坐标获取异常
            except exc.CoordinateError as e:
                metadata.wrong(str(e))
                continue
            cell_list = []
            try:
                for coord in coord_generator:
                    cell = self.__find_cell_or_create(coord, metadata.ref_class)
                    cell_list.append(cell)
            # 匹配坐标（集合）失败
            except exc.MisMatchingError as e:
                metadata.wrong(str(e))
                continue
            for cell in cell_list:
                cell.add_workload(metadata.issue, self.__jira_op, 1 / len(cell_list))
            metadata.success("Coordinate(s): " + ' & '.join([cell.coord_string for cell in cell_list]))
        assert len(self.__meta_datas) == sum(self.__num_of(res) for res in self.LoadResult)
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
        print("Exporting worklog table ...")
        concat_list = []
        for cell in self.__cells:
            concat_list.append(cell.get_worklog_table())
        table = pd.concat(concat_list).reset_index(drop=True)
        table.sort_values(by=list(table.columns[[0, 1, 4]]), inplace=True)
        print("Exporting worklog table completed.\n")
        return table

    def __num_of(self, res: LoadResult):
        return sum(map(lambda x: x.load_result is res, self.__meta_datas))

    def meta_data_loading_report(self, show_detail=True):
        print("Loading report: ")
        if show_detail:
            for res in self.LoadResult:
                print("\t%s(%s): " % (res.res_name, self.__num_of(res)))
                for metadata in self.__meta_datas:
                    if metadata.load_result is res:
                        print(utils.specific_length_string(metadata.issue.info_string), metadata.load_detail)
        print('Total: %d' % len(self.__meta_datas), end=', ')
        print(', '.join(map(lambda x: '%s: %d' % (x.res_name, self.__num_of(x)), self.LoadResult)) + '\n')

    def __build_matrix(self, ref_map: ReferenceMap, cell_property: Callable[[Cell], int | float]):
        array = np.zeros(ref_map.value_shape)
        for cell in self.__cells:
            if not cell.refer_from(ref_map):
                continue
            row_index, col_index = cell.coord_index
            array[row_index][col_index] = cell_property(cell)
        return array

    def __synthesize_sheet(self, ref_map: ReferenceMap, cell_property: Callable[[Cell], int | float], head='{}'):
        value_array = self.__build_matrix(ref_map, cell_property)
        worksheet = ref_map.value_array2sheet(value_array)
        worksheet.cell(1, 1).value = head.format(ref_map.worksheet_name)
        return worksheet

    def __downmix_sheet(self, ref_map: ReferenceMap, downmix_x: int | None, downmix_y: int | None,
                        cell_property: Callable[[Cell], int | float], head='{}'):
        value_array = self.__build_matrix(ref_map, cell_property)
        worksheet = ref_map.value_array2downmix_sheet(value_array, downmix_x, downmix_y)
        worksheet.cell(1, 1).value = (head.format(ref_map.worksheet_name)
                                      + ' downmix by (%s, %s)' % (downmix_x, downmix_y))
        return worksheet

    def export_matrix_workbook(self):
        unit = 'day'

        def cell_count(cell: Cell):
            return cell.num_issues

        def cell_workload(cell: Cell):
            return cell.cumulative_workload(unit=unit)

        count_head = r"Count of {}"
        workload_head = r"Cumulative workload of {}(person·%s)" % unit
        print("Exporting matrix workbook ...")
        workbook = Workbook()
        worksheet = workbook.active
        # 测试计数
        print("Building count matrix and synthesizing with ref_test style ...")
        test_count = self.__synthesize_sheet(self.__ref_test, cell_count, count_head)
        WorksheetProcessor.copy_into(test_count, worksheet)
        worksheet.title = 'Count of Test'
        # 测试计数压缩
        print("Building DOWNMIX count matrix base on ref_test ...")
        test_count_dm = self.__downmix_sheet(self.__ref_test, None, 1, cell_count, count_head)
        WorksheetProcessor.copy_into(test_count_dm, workbook.create_sheet('Count of Test(DOWNMIX)'))
        # 测试计时
        print("Building workload matrix and synthesizing with ref_test style ...")
        test_workload = self.__synthesize_sheet(self.__ref_test, cell_workload, workload_head)
        WorksheetProcessor.copy_into(test_workload, workbook.create_sheet('Time of Test'))
        # 测试计时压缩
        print("Building DOWNMIX workload matrix base on ref_test ...")
        test_workload_dm = self.__downmix_sheet(self.__ref_test, None, 1, cell_workload, workload_head)
        WorksheetProcessor.copy_into(test_workload_dm, workbook.create_sheet('Time of Test(DOWNMIX)'))
        # 管理计数
        print("Building count matrix and synthesizing with ref_manage style ...")
        manage_count = self.__synthesize_sheet(self.__ref_manage, cell_count, count_head)
        WorksheetProcessor.copy_into(manage_count, workbook.create_sheet('Count of Manage'))
        # 管理计数压缩
        pass
        # 管理计时
        print("Building workload matrix and synthesizing with ref_manage style ...")
        manage_workload = self.__synthesize_sheet(self.__ref_manage, cell_workload, workload_head)
        WorksheetProcessor.copy_into(manage_workload, workbook.create_sheet('Time of Manage'))
        print("Exporting matrix workbook completed\n")
        # 管理计时压缩
        pass
        return workbook
