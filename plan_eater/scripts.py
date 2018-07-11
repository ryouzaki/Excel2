from openpyxl import *
from openpyxl.formatting import Rule
from openpyxl.formatting.rule import *
from openpyxl.styles import Font, PatternFill, Border, Side, GradientFill
from openpyxl.styles.numbers import *
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles import Alignment
from openpyxl.worksheet.cell_range import CellRange
import re
import os
import datetime
from copy import copy


def create_new_json(plan):
    workbook = load_workbook(plan)
    mediaplan_sheet = workbook['MediaPlan']
    placement_dict = dict()
    weeknumber_set = set()
    for row in range(1,mediaplan_sheet.max_row):
        if re.search('\w\w_\w\w_\d\d\d', str(mediaplan_sheet.cell(row=row, column=1).value)):
            placement_dict.clear()
            weeknumber_set.clear()
            for column in range(2, mediaplan_sheet.max_column+1):
                if mediaplan_sheet.cell(row=get_fields_row(mediaplan_sheet), column=column).value is not None:
                    if re.search('\d{1,2}', str(mediaplan_sheet.cell(row=get_fields_row(mediaplan_sheet), column=column).value)):
                        if mediaplan_sheet.cell(row=row, column=column).value == 1:
                            weeknumber_set.add(int(mediaplan_sheet.cell(row=get_fields_row(mediaplan_sheet), column=column).value))
                        continue
                    placement_dict.update({mediaplan_sheet.cell(row=get_fields_row(mediaplan_sheet),
                                                                column=column).value: mediaplan_sheet.cell(
                        row=row, column=column).value})
            week_list = list()
            for weeknumber in sorted(weeknumber_set):
                week_list.append(
                    {'weeknumber': weeknumber, 'fact_budget': None, 'fact_impressions': None, 'fact_views': None,
                        'fact_clicks': None,
                        'fact_reach': None})
                placement_dict.update({'postclick': week_list})
            with open(str(os.path.dirname(plan)) + '\\' + str(
                    mediaplan_sheet.cell(row=row, column=1).value) + '.json', 'w') as outfile:
                json.dump(placement_dict, outfile)
    return


def update_json(plan):
    workbook = load_workbook(plan)
    mediaplan_sheet = workbook['MediaPlan']
    weeknumber_set = set()
    for row in range(1, mediaplan_sheet.max_row):
        if re.search('\w\w_\w\w_\d\d\d', str(mediaplan_sheet.cell(row=row, column=1).value)):
            with open(str(os.path.dirname(plan)) + '\\' + str(mediaplan_sheet.cell(row=row, column=1).value) + '.json',
                      'r') as infile:
                placement_dict = json.load(infile)
            week_list = placement_dict['postclick']
            for week in week_list:
                if week['weeknumber'] > datetime.datetime.today().isocalendar()[1]:
                    week_list.remove(week)
            for column in range(2, mediaplan_sheet.max_column + 1):
                if mediaplan_sheet.cell(row=get_fields_row(mediaplan_sheet), column=column).value is not None:
                    if re.search('\d{1,2}',
                                 str(mediaplan_sheet.cell(row=get_fields_row(mediaplan_sheet), column=column).value)):
                        if (mediaplan_sheet.cell(row=row, column=column).value == 1) and (
                                datetime.datetime.today().isocalendar()[1] <
                                (mediaplan_sheet.cell(row=get_fields_row(mediaplan_sheet),
                                                      column=column).value)):
                            weeknumber_set.add(mediaplan_sheet.cell(row=get_fields_row(mediaplan_sheet),
                                                                    column=column).value)
                            print (mediaplan_sheet.cell(row=get_fields_row(mediaplan_sheet),
                                                                    column=column).value)
                        continue
                    placement_dict.update({mediaplan_sheet.cell(row=get_fields_row(mediaplan_sheet),
                                                                column=column).value: mediaplan_sheet.cell(row=row,
                                                                                                           column=column).value})
            for weeknumber in weeknumber_set:
                week_list.append(
                    {'weeknumber': weeknumber, 'fact_budget': None, 'fact_impressions': None, 'fact_views': None,
                     'fact_clicks': None, 'fact_reach': None})
            placement_dict.update({'postclick': week_list})
            with open(str(os.path.dirname(plan)) + '\\' + str(mediaplan_sheet.cell(row=row, column=1).value) + '.json',
                      'w') as outfile:
                json.dump(placement_dict, outfile)


def parse_amnet():
    report = load_workbook(os.getcwd()+'\Reports\\Amnet\\Amnet'+str(datetime.datetime.today().isocalendar()[1])+'.xlsx')
    week_report_sheet = report[str(datetime.datetime.today().isocalendar()[1])]
    report_dict = dict()
    row = int()
    while row < week_report_sheet.max_row:
        row+=1
        if re.search('\w\w_\w\w_\d\d\d', str(week_report_sheet.cell(row=row, column=1).value)):
            placement_id = str(week_report_sheet.cell(row=row, column=1).value)
            print (str(week_report_sheet.cell(row=row, column=1).value))
            while week_report_sheet.cell(row=row+1, column=1).value != None:
                report_dict.update({str(week_report_sheet.cell(row=row+1, column=1).value): week_report_sheet.cell(row=row+1, column=2).value})
                row+=1
            print(report_dict)
            with open(str(os.getcwd()) + '\\MP\\' + placement_id + '.json', 'r+') as infile:
                placement_dict = json.load(infile)
                infile.seek(0)
                infile.truncate()
                week_list = placement_dict['postclick']
                for week_info in week_list:
                    if week_info['weeknumber'] == datetime.datetime.today().isocalendar()[1]:
                        print (week_info['weeknumber'])
                        print(report_dict)
                        for stat in week_info.keys():
                            week_info.update({stat:report_dict.get(stat)})
                            week_info.update({'weeknumber':datetime.datetime.today().isocalendar()[1]})
                        print (week_info)
                print (week_list)
                placement_dict.update({'postclick': week_list})
                json.dump(placement_dict, infile)
                report_dict.clear()
    return


def create_report(week):
    report = load_workbook(os.getcwd()+'\Reports\\'+ 'Template.xlsx')
    week_set = set()
    for jsonplan in os.listdir(os.getcwd() + '\\MP\\'):
        if jsonplan.endswith('.json'):
            with open(str(os.getcwd()) + '\\MP\\' + jsonplan, 'r') as infile:
                placement_dict = json.load(infile)
                for value in placement_dict['postclick']:
                    #if value['weeknumber'] <= datetime.datetime.today().isocalendar()[1]:
                        week_set.add(value['weeknumber'])
    for week in sorted(week_set):
        fields_row=1
        source = report.active
        target = report.copy_worksheet(source)
        first_stage_row= target.max_row+1
        medium = Side(border_style="medium", color="000000")
        borders = Border (top=medium, left=medium, right=medium, bottom=medium)
        while target.cell(row=fields_row,column=1).value != "fields":
            fields_row+=1
        for stage in ['Awareness','Consideration', 'Preference', 'Action', 'Loyalty']:
            if first_stage_row!=target.max_row+1:
                target.cell(row=first_stage_row, column = 2).alignment = Alignment(textRotation=90)
                target.merge_cells(start_row=first_stage_row, end_row=target.max_row, start_column=2, end_column=2)
                style_merged_cells(target,first_stage_row,target.max_row,2,2,border=borders, fill = GradientFill(stop=("ffff99", "ffff99")))
                first_stage_row=target.max_row+1
            for category in ['OLV', 'Programmatic', 'Social Media', 'SEA']:
                category_flag = True
                for jsonplan in os.listdir(os.getcwd() + '\\MP\\'):
                    if jsonplan.endswith('.json'):
                        with open(str(os.getcwd()) + '\\MP\\' + jsonplan, 'r') as infile:
                            placement_dict = json.load(infile)
                            if placement_dict["stage"]==stage and placement_dict["category"]==category:
                                last_column = 1
                                last_row = target.max_row+1
                                if category_flag:
                                    target.cell(row=last_row,column=last_column+1, value = stage)
                                    target.merge_cells(start_row=last_row, start_column=last_column+2,end_row=last_row,end_column=target.max_column-1)
                                    target.cell(row=last_row,column=last_column+2, value = category)
                                    style_merged_cells(target,last_row,last_row,last_column+2,target.max_column-1,border=borders,fill = GradientFill(stop=("ffff99", "ffff99")))
                                    category_flag = False
                                    last_row+=1
                                while (target.cell(row=fields_row, column = last_column).value) != "end":
                                    target.cell(row=last_row,column=last_column, value = placement_dict.get(str(target.cell(row=fields_row, column = last_column).value)))

                                    if target.cell(row=fields_row, column = last_column).value == "period":
                                        target.cell(row=last_row,column=last_column, value = str(len(placement_dict["postclick"])) + " weeks")

                                    if re.search('\d{1,2}', str(target.cell(row=fields_row, column = last_column).value)):
                                        for value in placement_dict['postclick']:
                                            if value['weeknumber'] == target.cell(row=fields_row, column = last_column).value:
                                                target.cell(row=last_row, column=last_column).fill = PatternFill(patternType='solid',fill_type='solid', fgColor = Color('3b3b3b'))

                                    if target.cell(row=fields_row, column = last_column).value == "plan_impressions":
                                        target.cell(row=last_row, column=last_column, value = (placement_dict.get("plan_impressions")))
                                        target.cell(row=last_row, column=last_column).number_format = '#,##0_-₽'
                                        target.cell(row=last_row, column=last_column).number_format = '#,##0_-'

                                    if target.cell(row=fields_row, column = last_column).value == "plan_reach":
                                        target.cell(row=last_row, column=last_column, value = (placement_dict.get("plan_reach")))
                                        target.cell(row=last_row, column=last_column).number_format = '#,##0_-'

                                    if target.cell(row=fields_row, column = last_column).value == "plan_clicks":
                                        target.cell(row=last_row, column=last_column, value = (placement_dict.get("plan_clicks")))
                                        target.cell(row=last_row, column=last_column).number_format = '#,##0_-'

                                    if target.cell(row=fields_row, column = last_column).value == "plan_views":
                                        target.cell(row=last_row, column=last_column, value = (placement_dict.get("plan_views")))
                                        target.cell(row=last_row, column=last_column).number_format = '#,##0_-'

                                    if target.cell(row=fields_row, column = last_column).value == "plan_budget":
                                        target.cell(row=last_row, column=last_column, value = (placement_dict.get("plan_budget")))
                                        target.cell(row=last_row, column=last_column).number_format = '#,##0_-₽'

                                    if target.cell(row=fields_row, column = last_column).value == "fact_impressions":
                                        for value in placement_dict['postclick']:
                                            if value['weeknumber'] == week:
                                                target.cell(row=last_row, column=last_column, value = value.get("fact_impressions"))
                                                target.cell(row=last_row, column=last_column).number_format = '#,##0_-'

                                    if target.cell(row=fields_row, column = last_column).value == "fact_clicks":
                                        for value in placement_dict['postclick']:
                                            if value['weeknumber'] == week:
                                                target.cell(row=last_row, column=last_column, value = value.get("fact_clicks"))
                                                target.cell(row=last_row, column=last_column).number_format = '#,##0_-'

                                    if target.cell(row=fields_row, column = last_column).value == "fact_budget":
                                        for value in placement_dict['postclick']:
                                            if value['weeknumber'] == week:
                                                target.cell(row=last_row, column=last_column, value = value.get("fact_budget"))
                                                target.cell(row=last_row, column=last_column).number_format = '#,##0_-₽'

                                    if target.cell(row=fields_row, column = last_column).value == "fact_reach":
                                        for value in placement_dict['postclick']:
                                            if value['weeknumber'] == week:
                                                target.cell(row=last_row, column=last_column, value = value.get("fact_reach"))
                                                target.cell(row=last_row, column=last_column).number_format = '#,##0_-'

                                    if target.cell(row=fields_row, column = last_column).value == "fact_views":
                                        for value in placement_dict['postclick']:
                                            if value['weeknumber'] == week:
                                                target.cell(row=last_row, column=last_column, value = value.get("fact_views"))
                                                target.cell(row=last_row, column=last_column).number_format = '#,##0_-'

                                    if target.cell(row=fields_row, column = last_column).value == "plan_cpm":
                                        target.cell(row=last_row, column=last_column, value = (placement_dict.get("plan_budget")*1000/placement_dict.get("plan_impressions")))
                                        target.cell(row=last_row, column=last_column).number_format = '#,##0_-₽'

                                    if target.cell(row=fields_row, column = last_column).value == "plan_cpt":
                                        target.cell(row=last_row, column=last_column, value = (placement_dict.get("plan_budget")*1000/placement_dict.get("plan_reach")))
                                        target.cell(row=last_row, column=last_column).number_format = '#,##0_-₽'

                                    if target.cell(row=fields_row, column = last_column).value == "plan_ctr":
                                        target.cell(row=last_row, column=last_column, value = (placement_dict.get("plan_clicks")/placement_dict.get("plan_impressions")))
                                        target.cell(row=last_row, column=last_column).number_format = FORMAT_PERCENTAGE_00

                                    if target.cell(row=fields_row, column = last_column).value == "plan_cpc":
                                        target.cell(row=last_row, column=last_column, value = (placement_dict.get("plan_budget")/placement_dict.get("plan_clicks")))
                                        target.cell(row=last_row, column=last_column).number_format = '#,##0.00_-₽'

                                    if target.cell(row=fields_row, column = last_column).value == "plan_vtr":
                                        target.cell(row=last_row, column=last_column, value = (placement_dict.get("plan_views")/placement_dict.get("plan_impressions")))
                                        target.cell(row=last_row, column=last_column).number_format = FORMAT_PERCENTAGE_00

                                    if target.cell(row=fields_row, column = last_column).value == "plan_cpv":
                                        target.cell(row=last_row, column=last_column, value = (placement_dict.get("plan_budget")/placement_dict.get("plan_views")))
                                        target.cell(row=last_row, column=last_column).number_format = '#,##0.00_-₽'

                                    if target.cell(row=fields_row, column = last_column).value == "fact_cpm":
                                        for value in placement_dict['postclick']:
                                            if value['weeknumber'] == week:
                                                print (week)
                                                target.cell(row=last_row, column=last_column, value = (value.get("fact_budget")*1000/value.get("fact_impressions")))
                                                target.cell(row=last_row, column=last_column).number_format = '#,##0_-₽'

                                    if target.cell(row=fields_row, column = last_column).value == "fact_cpt":
                                        for value in placement_dict['postclick']:
                                            if value['weeknumber'] == week:
                                                target.cell(row=last_row, column=last_column, value = (value.get("fact_budget")*1000/value.get("fact_reach")))
                                                target.cell(row=last_row, column=last_column).number_format = '#,##0_-₽'

                                    if target.cell(row=fields_row, column = last_column).value == "fact_ctr":
                                        for value in placement_dict['postclick']:
                                            if value['weeknumber'] == week:
                                                target.cell(row=last_row, column=last_column, value = (value.get("fact_clicks")/value.get("fact_impressions")))
                                                target.cell(row=last_row, column=last_column).number_format = FORMAT_PERCENTAGE_00

                                    if target.cell(row=fields_row, column = last_column).value == "fact_cpc":
                                        for value in placement_dict['postclick']:
                                            if value['weeknumber'] == week:
                                                target.cell(row=last_row, column=last_column, value = (value.get("fact_budget")/value.get("fact_clicks")))
                                                target.cell(row=last_row, column=last_column).number_format = '#,##0.00_-₽'

                                    if target.cell(row=fields_row, column = last_column).value == "fact_vtr":
                                        for value in placement_dict['postclick']:
                                            if value['weeknumber'] == week:
                                                target.cell(row=last_row, column=last_column, value = (value.get("fact_views")/value.get("fact_impressions")))
                                                target.cell(row=last_row, column=last_column).number_format = FORMAT_PERCENTAGE_00

                                    if target.cell(row=fields_row, column = last_column).value == "fact_cpv":
                                        for value in placement_dict['postclick']:
                                            if value['weeknumber'] == week:
                                                target.cell(row=last_row, column=last_column, value = (value.get("fact_budget")/value.get("fact_views")))
                                                target.cell(row=last_row, column=last_column).number_format = '#,##0.00_-₽'
                                    target.cell(row=last_row,column=last_column).border = copy(target.cell(row=fields_row, column = last_column).border)
                                    last_column+=1

        d = str(datetime.datetime.now().year) + '-W' + str(week)
        target.title = datetime.datetime.strftime(datetime.datetime.strptime(d + '-1', "%Y-W%W-%w"),"%d-%b-%Y") + " -- " + datetime.datetime.strftime(datetime.datetime.strptime(d + '-1', "%Y-W%W-%w")+datetime.timedelta(days=6),"%d-%b-%Y")
    report.save(os.getcwd()+'\Reports\\Client\\'+ 'Report_week' +str(datetime.datetime.today().isocalendar()[1])+'.xlsx')
    return


def style_merged_cells(ws, min_row, max_row, min_col, max_col, border=Border(), fill=None, font=None, alignment=None):
    top = Border(top=border.top)
    left = Border(left=border.left)
    right = Border(right=border.right)
    bottom = Border(bottom=border.bottom)
    first_cell = ws.cell(min_row,min_col)
    if alignment:
        ws.merge_cells(start_row=min_row,end_row=max_row,start_column=min_col,end_column=max_col)
        first_cell.alignment = alignment
    if font:
        first_cell.font = font
    for rows in ws.iter_rows(min_col=min_col, max_col=max_col, min_row=min_row,max_row=max_row):
        for cell in rows:
            cell.border = cell.border + top
    for rows in ws.iter_rows(min_col=min_col, max_col=max_col,min_row=min_row, max_row=max_row):
        for cell in rows:
            cell.border = cell.border + bottom
    for columns in ws.iter_cols(min_col=min_col, max_col=max_col,min_row=min_row, max_row=max_row):
        for cell in columns:
            cell.border = cell.border + left
    for columns in ws.iter_cols(min_col=min_col, max_col=max_col,min_row=min_row, max_row=max_row):
        for cell in columns:
            cell.border = cell.border + right
            if fill:
                cell.fill=fill


def get_fields_row(worksheet):
    fields_row=1
    while worksheet.cell(row=fields_row,column=1).value != "fields":
        fields_row+=1
    return fields_row




