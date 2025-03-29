from models import JIRALogin, IssueList, ConcatFilter, JIRAOperator, Matrix, WorksheetShell
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


def export_worklog_workbook(wm: Matrix):
    df = wm.export_worklog_table()
    workbook = Workbook()
    worksheet = workbook.active
    for row_content in dataframe_to_rows(df, index=False, header=True):
        worksheet.append(row_content)
    WorksheetShell(worksheet).batch_set_column_width({
        'A': 60,
        'B': 40,
        'C': 40,
        'D': 40,
    })
    filename = 'Export Worklog.xlsx'
    workbook.save(filename)
    print('Worklog table save as %s.' % filename)


def export_matrix_workbook(wm: Matrix):
    workbook = wm.export_matrix_workbook()
    filename = 'MatrixAggregation.xlsx'
    workbook.save(filename)
    print('Matrix Aggregation save as: %s' % filename)


if __name__ == '__main__':
    jira_agent = JIRALogin.used_token(r'access_token.txt')
    issue_list = IssueList(jira_agent.get_fields())
    # issue_list.import_issues(jira_agent.search_by_jql_filter(BaseFilter.ALL_TASK_LIKE))
    # issue_list.import_issues(jira_agent.search_by_jql_filter(BaseFilter.ALL_TASK_LIKE_GOOGLE))
    issue_list.import_issues(jira_agent.search_by_jql_filter(ConcatFilter.ALL_TASK_LIKE))
    jira_op = JIRAOperator(jira_agent)
    workload_matrix = Matrix(issue_list, jira_op, '2025年标准工时时间表.xlsx')
    load_report = workload_matrix.meta_data_loading_report()
    load_report.to_excel('LoadingReport.xlsx', header=True, index=False)
    export_worklog_workbook(workload_matrix)
    print(jira_op.call_num_log, '\n')
    export_matrix_workbook(workload_matrix)
