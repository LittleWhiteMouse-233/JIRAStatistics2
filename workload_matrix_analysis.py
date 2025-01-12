from models import JIRALogin, IssueList, BaseFilter, JIRAOperator, Matrix, WorksheetProcessor
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


def export_worklog_workbook(wm: Matrix):
    df = wm.export_worklog_table()
    workbook = Workbook()
    worksheet = workbook.active
    for row_content in dataframe_to_rows(df, index=False, header=True):
        worksheet.append(row_content)
    WorksheetProcessor(worksheet).batch_set_column_width({
        'A': 60,
        'B': 40,
        'C': 40,
        'D': 40,
    })
    filename = 'Export Worklog.xlsx'
    workbook.save(filename)
    print('Worklog table save as %s.' % filename)


if __name__ == '__main__':
    jira_agent = JIRALogin.used_token(r'access_token.txt')
    issue_list = IssueList(jira_agent.get_fields())
    issue_list.import_issues(jira_agent.search_by_jql_filter(BaseFilter.ALL_TASK_LIKE))
    jira_op = JIRAOperator(jira_agent)
    workload_matrix = Matrix(issue_list, jira_op, '2024年标准工时时间表.xlsx')
    workload_matrix.meta_data_loading_report()
    export_worklog_workbook(workload_matrix)
    print(jira_op.call_num_log, '\n')
    wb = workload_matrix.export_matrix_workbook()
    wb.save('Matrix.xlsx')
