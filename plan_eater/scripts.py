from openpyxl import *
import re
import os
import datetime


def create_new_json(plan):
    workbook = load_workbook(plan)
    mediaplan_sheet = workbook['MediaPlan']
    last_row = 1
    placement_dict = dict()
    weeknumber_set = set()
    while mediaplan_sheet.cell(row=last_row, column=1).value != "borders":
        if re.search('\w\w_\w\w_\d\d\d', str(mediaplan_sheet.cell(row=last_row, column=1).value)):
            placement_dict.clear()
            weeknumber_set.clear()
            last_column = 2
            while mediaplan_sheet.cell(row=4, column=last_column).value != "end":
                placement_dict.update({mediaplan_sheet.cell(row=4, column=last_column).value: mediaplan_sheet.cell(
                    row=last_row, column=last_column).value})
                if (9 < last_column < 375 and mediaplan_sheet.cell(row=last_row, column=last_column).value == 1):
                    weeknumber_set.add(mediaplan_sheet.cell(row=10, column=last_column).value.isocalendar()[1])
                last_column += 1
            week_list = list()
            for weeknumber in weeknumber_set:
                week_list.append(
                {'weeknumber': weeknumber, 'budget': None, 'impressions': None, 'views': None, 'clicks': None,
                     'reach': None})
            placement_dict.update({'postclick': week_list})
            with open(str(os.path.dirname(plan)) + '\\' + str(mediaplan_sheet.cell(row=last_row, column=1).value) + '.json', 'w') as outfile:
                json.dump(placement_dict, outfile)
        last_row += 1
    return


def update_json(plan):
    workbook = load_workbook(plan)
    mediaplan_sheet = workbook['MediaPlan']
    last_row = 1
    weeknumber_set = set()
    while mediaplan_sheet.cell(row=last_row, column=1).value != "borders":
        if re.search('\w\w_\w\w_\d\d\d', str(mediaplan_sheet.cell(row=last_row, column=1).value)):
            with open(str(os.path.dirname(plan)) + '\\' + str(mediaplan_sheet.cell(row=last_row, column=1).value) + '.json', 'r') as infile:
                placement_dict = json.load(infile)
            week_list = placement_dict['postclick']
            j=1
            for week in week_list:
                if week['weeknumber'] > datetime.datetime.today().isocalendar()[1]:
                    week_list.remove(week)
            last_column = 1
            while mediaplan_sheet.cell(row=4, column=last_column).value != "end":
                placement_dict.update({mediaplan_sheet.cell(row=4, column=last_column).value: mediaplan_sheet.cell(row=last_row, column=last_column).value})
                if ((last_column > 9 and last_column < 375) and (mediaplan_sheet.cell(row=last_row, column=last_column).value == 1) and (datetime.datetime.today().isocalendar()[1] < (mediaplan_sheet.cell(row=10, column=last_column).value).isocalendar()[1])):
                        weeknumber_set.add((mediaplan_sheet.cell(row=10, column=last_column).value).isocalendar()[1])
                last_column += 1
            print (week_list)
            for weeknumber in weeknumber_set:

                week_list.append({'weeknumber': weeknumber, 'budget': None, 'impressions': None, 'views': None, 'clicks': None,'reach': None})
            placement_dict.update({'postclick': week_list})
            with open(str(os.path.dirname(plan)) + '\\' + str(mediaplan_sheet.cell(row=last_row, column=1).value) + '.json', 'w') as outfile:
                json.dump(placement_dict, outfile)
        last_row += 1


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
                report_dict.update({week_report_sheet.cell(row=row+1, column=1).value: week_report_sheet.cell(row=row+1, column=2).value})
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
    report = Workbook()
    week_set = set()
    for jsonplan in os.listdir(os.getcwd() + '\\MP\\'):
        if jsonplan.endswith('.json'):
            with open(str(os.getcwd()) + '\\MP\\' + jsonplan, 'r') as infile:
                placement_dict = json.load(infile)
                for value in placement_dict['postclick']:
                    if value['weeknumber'] < datetime.datetime.today().isocalendar()[1]:
                        week_set.add(value['weeknumber'])
    for week in sorted(week_set):
        report.create_sheet(str(week))

    #insert plan data and periods of placement
    for sheet in report.worksheets:
        sheet.merge_cells(start_row = 7, start_column=1, end_row=9,end_column=1)
        for jsonplan in os.listdir(os.getcwd() + '\\MP\\'):
            if jsonplan.endswith('.json'):
                with open(str(os.getcwd()) + '\\MP\\' + jsonplan, 'r') as infile:
                    placement_dict = json.load(infile)
                    for value in placement_dict['postclick']:
                        if str(value['weeknumber']) == sheet.title:
                            sheet.cell(row=sheet.max_row, column=sheet.max_column+1, value = jsonplan)
                            sheet.cell(row=sheet.max_row+1, column=sheet.max_column, value = placement_dict['kpi'])
                            print (sheet.max_row)
    report.save(os.getcwd()+'\\Reports\\Client\\'+str(datetime.datetime.today().isocalendar()[1])+'.xlsx')
    return
