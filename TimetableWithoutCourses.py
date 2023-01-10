from traceback import print_exc


try:

    # %% [markdown]
    # os.listdir 遍历文件夹
    # openpyxl.Workbook 建立excel工作簿、表
    # openpyxl.styles excel样式
    # pdfplumber 读取、解析pdf内容
    # re 正则

    # %%
    import sys
    import os
    from openpyxl import Workbook
    from openpyxl.styles import *
    import pdfplumber
    import re

    # %% [markdown]
    # 基础变量

    # %%
    week_name = ['星期一', '星期二', '星期三', '星期四', '星期五']

    # 改变工作目录到脚本所在目录
    os.chdir(os.path.split(os.path.abspath(sys.argv[0]))[0])

    timetables_dir = '课表'
    departments_dir = os.listdir(timetables_dir)
    departments_dir.sort()
    departments = departments_dir.copy()
    departments_num = len(departments)

    for i, department in enumerate(departments):
        for j in range(len(department)):
            if not department[j].isdigit():
                departments[i] = department[j:]
                break

    for i in range(departments_num):
        departments_dir[i] = os.path.join(timetables_dir, departments_dir[i])

    # 默认总周数weeks_num=16
    weeks_num = 16
    # 获取总周数
    for filepath, dirnames, filenames in os.walk(timetables_dir):
        for filename in filenames:
            with pdfplumber.open(os.path.join(filepath, filename)) as pdf:
                table: list[list[str]] = pdf.pages[0].extract_table()
                for row in table:
                    for info in row:
                        if info:
                            if info.count('18周'):
                                weeks_num = 18
                                break
                            elif info.count('17周'):
                                weeks_num = 17
                                break
                            elif info.count('16周'):
                                weeks_num = 16
                                break
                    else:
                        continue
                    break
                else:
                    continue
                break
        else:
            continue
        break

    re_class_time = re.compile(r'\(([1-9]|1[0-2])-([1-9]|1[0-2])节\)(([1-9]|1[0-' + str(
        weeks_num % 10) + r'])(-([1-9]|1[0-' + str(weeks_num % 10) + r']))?周(\((单|双)\))?,?)+')

    # %% [markdown]
    # excel表行列变量

    # %%
    start_row_morning = 3
    start_row_afternoon = start_row_morning + 3 * departments_num + 1
    start_row_evening = start_row_afternoon + 3 * departments_num + 1

    start_row_1, start_row_3, start_row_5 = range(
        start_row_morning, start_row_afternoon - 2, departments_num)
    start_row_6, start_row_8, start_row_9 = range(
        start_row_afternoon, start_row_evening - 2, departments_num)
    start_row_10 = start_row_evening
    start_row_12 = start_row_10 + departments_num

    start_row_class_table = [2, 4, 6, 7, 9, 10, 11, 13]
    start_row_class_output = [start_row_1, start_row_3, start_row_5, start_row_6,
                              start_row_8, start_row_9, start_row_10, start_row_12]
    class2table = [0, 0, 1, 1, 2, 3, 3, 4, 5, 6, 6, 7]

    # %% [markdown]
    # 新建工作簿wb 工作表ws

    # %%
    wb = Workbook()
    ws = wb.active

    # %% [markdown]
    # 设置ws样式 框架

    # %%
    grey = Color('FFCCCCCC')
    black = Color('FF000000')

    alignment = Alignment(horizontal='center',
                          vertical='center', wrap_text=True)
    font = Font(size=12)
    side = Side(border_style='medium', color=black)
    border = Border(side, side, side, side)

    # 所有单元格 设为字符串 居中 自动换行 字体 边框
    for i in range(1, start_row_12 + departments_num):
        for j in range(1, 9):
            ws.cell(i, j).value = ''
            ws.cell(i, j).alignment = alignment
            ws.cell(i, j).font = font
            ws.cell(i, j).border = border

    ws.merge_cells(None, 1, 1, 1, 8)
    ws.cell(1, 1).value = '无课表'
    ws.cell(1, 1).font = Font(size=14)
    ws.row_dimensions[1].height = 43

    ws.merge_cells(None, 2, 1, 2, 2)
    ws.cell(2, 1).value = '时间'
    ws.cell(2, 3).value = '部门'
    for i in range(5):
        ws.cell(2, i + 4).value = week_name[i]
    ws.row_dimensions[2].height = 36
    for i in range(3):
        ws.column_dimensions[chr(ord('A') + i)].width = 11
    for i in range(3, 8):
        ws.column_dimensions[chr(ord('A') + i)].width = 33

    ws.merge_cells(None, start_row_morning, 1, start_row_afternoon - 2, 1)
    ws.cell(start_row_morning, 1).value = '上午'
    ws.merge_cells(None, start_row_afternoon, 1, start_row_evening - 2, 1)
    ws.cell(start_row_afternoon, 1).value = '下午'
    ws.merge_cells(None, start_row_evening, 1,
                   start_row_evening + 2 * departments_num - 1, 1)
    ws.cell(start_row_evening, 1).value = '晚上'

    # 上下晚中间的空行
    ws.merge_cells(None, start_row_afternoon - 1,
                   1, start_row_afternoon - 1, 8)
    ws.merge_cells(None, start_row_evening - 1, 1, start_row_evening - 1, 8)

    ws.merge_cells(None, start_row_1, 2, start_row_3 - 1, 2)
    ws.cell(start_row_1, 2).value = '1-2'
    ws.merge_cells(None, start_row_3, 2, start_row_5 - 1, 2)
    ws.cell(start_row_3, 2).value = '3-4'
    ws.merge_cells(None, start_row_5, 2, start_row_6 - 2, 2)
    ws.cell(start_row_5, 2).value = '5'
    ws.merge_cells(None, start_row_6, 2, start_row_8 - 1, 2)
    ws.cell(start_row_6, 2).value = '6-7'
    ws.merge_cells(None, start_row_8, 2, start_row_9 - 1, 2)
    ws.cell(start_row_8, 2).value = '8'
    ws.merge_cells(None, start_row_9, 2, start_row_10 - 2, 2)
    ws.cell(start_row_9, 2).value = '9'
    ws.merge_cells(None, start_row_10, 2, start_row_12 - 1, 2)
    ws.cell(start_row_10, 2).value = '10-11'
    ws.merge_cells(None, start_row_12, 2,
                   start_row_12 + departments_num - 1, 2)
    ws.cell(start_row_12, 2).value = '12'

    # 部门列
    for i in start_row_class_output:
        for j in range(departments_num):
            ws.cell(i + j, 3).value = departments[j]

    # %% [markdown]
    # 读取pdf课表信息 建立无课表

    # %%
    for i in range(departments_num):
        print(departments[i])
        timetables = os.listdir(departments_dir[i])
        if not timetables is None:
            timetables.sort()
            for file in timetables:
                with pdfplumber.open(os.path.join(departments_dir[i], file)) as pdf:
                    # 将多页的表格合并成一个表格
                    tables = [page.extract_table()
                              for page in pdf.pages if not page.extract_table() is None]
                    name = tables[0][0][0].split('课表')[0]
                    print('\t', name)
                    for j in range(len(tables) - 1):
                        # 后面一页还是课表内容 不是“其他课程...”、“：讲课学时...”行
                        if len(tables[j + 1][0]) > 1:
                            if (tables[j][-1][1] == '' or tables[j + 1][0][1] == ''):
                                for col in range(len(tables[j+1][0])):
                                    if (tables[j + 1][0][col] != ''):
                                        row_off = 0
                                        while tables[j][-1 - row_off][col] is None:
                                            row_off += 1
                                        tables[j][-1 -
                                                  row_off][col] += tables[j + 1][0][col]
                                tables[j + 1] = tables[j + 1][1:]
                    table: list[list[str]] = []
                    for t in tables:
                        table += t

                    # 建立个人无课表
                    table_without_courses = [
                        [set(range(1, weeks_num + 1)) for _ in range(5)] for _ in range(8)]
                    for k in range(5):
                        for j in range(8):
                            if table[start_row_class_table[j]][k + 2]:
                                table[start_row_class_table[j]][k + 2] = ''.join(
                                    table[start_row_class_table[j]][k + 2].split('\n'))
                                # class_times = re_class_time.findall(table[start_row_class_table[j]][k + 2])
                                class_times = re_class_time.finditer(
                                    table[start_row_class_table[j]][k + 2])
                                for class_time in class_times:
                                    class_time = class_time.group().split('节)')
                                    start_class, end_class = map(
                                        int, class_time[0][1:].split('-'))
                                    start_end_weeks = class_time[1].split(',')
                                    for start_end_week in start_end_weeks:
                                        start_end_week = start_end_week.split(
                                            '周')
                                        mark = start_end_week[1]
                                        start_end_week = start_end_week[0].split(
                                            '-')
                                        if len(start_end_week) == 1:
                                            start_week = end_week = int(
                                                start_end_week[0])
                                        else:
                                            start_week, end_week = map(
                                                int, start_end_week)
                                        if mark != '':
                                            mark = mark[1]
                                            if mark == '单':
                                                if not start_week & 1:
                                                    start_week += 1
                                                if not end_week & 1:
                                                    end_week -= 1
                                            elif mark == '双':
                                                if start_week & 1:
                                                    start_week += 1
                                                if end_week & 1:
                                                    end_week -= 1
                                        for l in range(class2table[start_class - 1], class2table[end_class - 1] + 1):
                                            table_without_courses[l][k] -= set(
                                                range(start_week, end_week + 1, 2 if mark else 1))

                    # 将个人无课表写入ws中
                    for k in range(5):
                        for j in range(8):
                            cur_table_without_courses = sorted(
                                list(table_without_courses[j][k]))
                            if cur_table_without_courses:
                                cell = ws.cell(
                                    start_row_class_output[j] + i, k + 4)
                                cell.value += '，' + name
                                if len(cur_table_without_courses) < weeks_num:
                                    value = ''
                                    start_week = end_week = cur_table_without_courses[0]
                                    for l in range(1, len(cur_table_without_courses)):
                                        if cur_table_without_courses[l] == cur_table_without_courses[l - 1] + 1:
                                            end_week = cur_table_without_courses[l]
                                        else:
                                            value += str(start_week)
                                            if end_week > start_week:
                                                if end_week > start_week+1:
                                                    value += '-'
                                                else:
                                                    value += '、'
                                                value += str(end_week)
                                            value += '、'
                                            start_week = end_week = cur_table_without_courses[l]
                                    value += str(start_week)
                                    if cur_table_without_courses[-1] > start_week:
                                        value += '-' + \
                                            str(cur_table_without_courses[l])
                                    cell.value += value
        #         break
        # break

    # 去除开头'，'
    for i in range(3, start_row_12 + departments_num):
        for j in range(4, 9):
            cell = ws.cell(i, j)
            if cell.value and cell.value[0] == '，':
                cell.value = cell.value[1:]

    # %%
    wb.save('无课表.xlsx')

    # %%
    input('按回车键退出')

except Exception:
    print_exc()
    print('！！！出错了，请反馈！！！\n' * 3 + 'qq 1932232849 或者github issues')
    input('按回车键退出')
