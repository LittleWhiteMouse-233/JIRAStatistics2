import re
from typing import Callable
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Border, Side, PatternFill
from copy import copy


class WorksheetProcessor:
    def __init__(self, worksheet: Worksheet):
        self.__ws = worksheet

    @property
    def worksheet(self):
        return self.__ws

    @property
    def max_row(self):
        return self.__ws.max_row

    @property
    def max_col(self):
        return self.__ws.max_column

    @staticmethod
    def __alpha2num(alpha: str):
        if not alpha.isalpha():
            raise ValueError("The parameter \'alpha\' string has non-alphabetic characters: " + str(alpha))
        num = 0
        for i, c in enumerate(alpha[::-1]):
            num += (ord(c.upper()) - ord('A') + 1) * pow(26, i)
        return num

    @staticmethod
    def __num2alpha(num: int):
        if num <= 0:
            raise ValueError("The parameter \'decimal_num\'(" + str(num) + ") cannot be less than 0.")
        alpha = ''
        while True:
            num, mod = divmod(num, 26)
            alpha = chr(mod + ord('A') - 1) + alpha
            if num == 0:
                break
        return alpha

    # 数字化列名称 & 超限检查
    def __activate_col(self, col: int | str):
        if type(col) is str:
            act_col = self.__alpha2num(col)
        else:
            act_col: int = col
        if type(act_col) is not int:
            raise ValueError("Invalid type of act_col: %s, excepted int or str" % type(act_col))
        if not 1 <= act_col <= self.max_col:
            raise ValueError("The act_col(%d) is out of range(%d:%d)" % (act_col, 1, self.max_col))
        return act_col

    # 数字化列名称列表 & 超限检查
    def __activate_col_list(self, col_list: list[int | str]):
        return sorted(list(set(map(lambda col: self.__activate_col(col), col_list))))

    @classmethod
    def __point_str2int(cls, point: str):
        return int(re.search(r'\d+', point).group()), cls.__alpha2num(re.search(r'[A-Z]+', point).group())

    @classmethod
    def __point_int2str(cls, point_row: int, point_col: int):
        return cls.__num2alpha(point_col) + str(point_row)

    # 数字化行列范围 & 超限检查
    def __activate_scope(self, scope_str: str):
        if not scope_str:
            return (1, self.max_row), (1, self.max_col)
        if ':' not in scope_str:
            raise ValueError("Invalid scope string: " + str(scope_str))
        begin, end = scope_str.upper().split(':')
        begin_row, begin_col = self.__point_str2int(begin)
        end_row, end_col = self.__point_str2int(end)

        def sort_xy(x: int, y: int):
            if x <= y:
                return x, y
            else:
                return y, x

        row_scope = sort_xy(begin_row, end_row)
        col_scope = sort_xy(begin_col, end_col)
        if not 1 <= row_scope[0] <= row_scope[1] <= self.max_row:
            raise ValueError("The row(%s) is out of range(%d:%d)" % (str(row_scope), 1, self.max_row))
        if not 1 <= col_scope[0] <= col_scope[1] <= self.max_col:
            raise ValueError("The col(%s) is out of range(%d:%d)" % (str(col_scope), 1, self.max_col))
        return row_scope, col_scope

    def __merge_cells_vertical(self, row_begin: int, row_end: int, col_begin: int, col_end: int, mode: str):
        excepted_mode = ['nan', 'same', 'all']
        if mode not in excepted_mode:
            raise ValueError("Merge mode should be: " + str(excepted_mode) + ", but given: " + str(mode))
        for col in range(col_begin, col_end + 1):
            start = row_begin
            end = row_begin
            last_str = ''
            for row in range(row_begin, row_end + 2):
                if row <= row_end:
                    cell_content = self.__ws.cell(row, col).value
                else:
                    cell_content = ''
                if (row > row_end
                        or (mode == 'nan' and type(cell_content) is str)
                        or (mode == 'same' and cell_content != last_str)
                        or (mode == 'all' and type(cell_content) is str and cell_content != last_str)):
                    if row_begin <= start < end <= row_end:
                        self.__ws.merge_cells(':'.join([self.__point_int2str(start, col),
                                                        self.__point_int2str(end, col)]))
                    start = row
                    if mode in ['same', 'all']:
                        last_str = cell_content
                end = row

    # 批量设置纵向单元格合并
    def batch_merge_cells_vertical(self, *, scope='', col_list: list = None, mode='all'):
        if col_list:
            act_col_list = self.__activate_col_list(col_list)
            for col in act_col_list:
                self.__merge_cells_vertical(1, self.max_row, col, col, mode=mode)
        else:
            row_scope, col_scope = self.__activate_scope(scope)
            self.__merge_cells_vertical(*row_scope, *col_scope, mode=mode)

    # 复制纵向单元格合并
    def copy_merge_cells_vertical(self, refer_col: int | str, target_col: int | str | list[int | str]):
        act_refer_col = self.__activate_col(refer_col)
        if type(target_col) is list:
            act_target_col = self.__activate_col_list(target_col)
        else:
            act_target_col = [self.__activate_col(target_col)]
        merge_list = []
        for merge_area in self.__ws.merged_cells:
            if merge_area.min_col == merge_area.max_col == act_refer_col:
                merge_list.append((merge_area.min_row, merge_area.max_row))
        for col in act_target_col:
            for begin, end in merge_list:
                self.__ws.merge_cells(':'.join([self.__point_int2str(begin, col),
                                                self.__point_int2str(end, col)]))

    # 批量设置列宽
    def batch_set_column_width(self, width_dict: dict[int | str, int]):
        for col, width in width_dict.items():
            if not isinstance(width, (int, float)):
                raise ValueError("Invalid width type: %s." % type(width))
            if width <= 0:
                raise ValueError("The column width should be positive, but given: %s" % width)
            act_col = self.__activate_col(col)
            self.__ws.column_dimensions[self.__num2alpha(act_col)].width = width
        return self.__ws

    # 解除合并单元格并填充
    def unmerge_cells_and_fill(self, fill=True):
        scope_list = []
        for merge_area in self.__ws.merged_cells:
            scope_list.append((merge_area.min_row, merge_area.min_col, merge_area.max_row, merge_area.max_col))
        for min_row, min_col, max_row, max_col in scope_list:
            self.__ws.unmerge_cells(start_row=min_row, start_column=min_col, end_row=max_row, end_column=max_col)
            if fill:
                for i in range(min_row, max_row + 1):
                    for j in range(min_col, max_col + 1):
                        self.__ws.cell(i, j).value = self.__ws.cell(min_row, min_col).value
        return self.__ws

    # 批量设置单元格
    def batch_set(self, func: Callable, scope='', col_list: list = None, **kwargs):
        if col_list:
            col_range = self.__activate_col_list(col_list)
            row_range = range(1, self.max_row + 1)
        else:
            row_scope, col_scope = self.__activate_scope(scope)
            row_range = range(row_scope[0], row_scope[1] + 1)
            col_range = range(col_scope[0], col_scope[1] + 1)
        for i in row_range:
            for j in col_range:
                func(i, j, **kwargs)

    # 设置单元格文本对齐
    def setting_text_alignment(self, row, col, horizontal='left', vertical='center'):
        excepted_h = ['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed']
        excepted_v = ['top', 'center', 'bottom', 'justify', 'distributed']
        if horizontal not in excepted_h:
            raise ValueError("Invalid horizontal: " + str(horizontal) + ', excepted input: ' + str(excepted_h))
        if vertical not in excepted_v:
            raise ValueError("Invalid vertical: " + str(vertical) + ', excepted input: ' + str(excepted_v))
        align = self.__ws.cell(row, col).alignment
        if align.horizontal != horizontal or align.vertical != vertical:
            align_new = copy(align)
            if align.horizontal != horizontal:
                align_new.horizontal = horizontal
            if align.vertical != vertical:
                align_new.vertical = vertical
            self.__ws.cell(row, col).alignment = align_new

    # 设置单元格文本自动换行
    def setting_word_wrap(self, row, col):
        align = self.__ws.cell(row, col).alignment
        if not align.wrapText:
            align_new = copy(align)
            align_new.wrapText = True
            self.__ws.cell(row, col).alignment = align_new

    # 设置单元格边框
    def setting_cell_border(self, row, col, border_style='thin'):
        excepted_style = ["dashDot", "dashDotDot", "dashed", "dotted", "double", "hair", "medium", "mediumDashDot",
                          "mediumDashDotDot", "mediumDashed", "slantDashDot", "thick", "thin", "none"]
        if border_style not in excepted_style:
            raise ValueError("Invalid border_style: %s, excepted input: %s" % (border_style, excepted_style))
        self.__ws.cell(row, col).border = Border(left=Side(border_style=border_style, color='FF000000'),
                                                 right=Side(border_style=border_style, color='FF000000'),
                                                 top=Side(border_style=border_style, color='FF000000'),
                                                 bottom=Side(border_style=border_style, color='FF000000'))

    # 设置单元格颜色
    def setting_fill_color(self, row, col, re_pattern: re.Pattern, fill_type='solid'):
        excepted_type = ["solid", "darkDown", "darkGray", "darkGrid", "darkHorizontal", "darkTrellis", "darkUp",
                         "darkVertical", "gray0625", "gray125", "lightDown", "lightGray", "lightGrid",
                         "lightHorizontal", "lightTrellis", "lightUp", "lightVertical", "mediumGray", "none"]
        if fill_type not in excepted_type:
            raise ValueError("Invalid fill_type: %s, excepted input: %s" % (fill_type, excepted_type))
        type_index = excepted_type.index(fill_type)
        content = self.__ws.cell(row, col).value
        color = re_pattern.search(content)
        if color:
            self.__ws.cell(row, col).fill = PatternFill(patternType=excepted_type[type_index],
                                                        fgColor=color.groups()[1])
            self.__ws.cell(row, col).value = re_pattern.sub('', content)

    # 设置单元格文本字体
    def setting_basic_font(self, row, col, name: str = None, size: int = None, bold: bool = None, color: str = None,
                           italic: bool = None, strike: bool = None):
        font_new = copy(self.__ws.cell(row, col).font)
        if name:
            font_new.name = name
        if size:
            font_new.size = size
        if bold:
            font_new.bold = bold
        if color:
            font_new.color = color
        if italic:
            font_new.italic = italic
        if strike:
            font_new.strike = strike
        self.__ws.cell(row, col).font = font_new
