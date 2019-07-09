'''This programme to calculate time difference between reported milestones

input documents:
There quarters master information, typically:
1) latest quarter data
2) last quarter data
3) year ago quarter data

output document:
Depending on instructions -
1) excel workbook with all project information

To operate the programme you should:
1) Enter the file paths to three quarters master data,
2) Specify the date from which to filter milestones,
3) Specify the specify the name and location of the output wb - consider whether to return one wb or multiple ones

'''

import datetime
from bcompiler.utils import project_data_from_master
from openpyxl import Workbook
from openpyxl.styles import Font



'''Function to filter out ALL milestone data'''
def all_milestone_data_bulk(project_list, master_data):
    upper_dict = {}

    for name in project_list:
        try:
            p_data = master_data[name]
            lower_dict = {}
            for i in range(1, 50):
                try:
                    try:
                        lower_dict[p_data['Approval MM' + str(i)]] = p_data['Approval MM' + str(i) + ' Forecast / Actual']
                    except KeyError:
                        lower_dict[p_data['Approval MM' + str(i)]] = p_data['Approval MM' + str(i) + ' Forecast - Actual']

                    lower_dict[p_data['Assurance MM' + str(i)]] = p_data['Assurance MM' + str(i) + ' Forecast - Actual']
                except KeyError:
                    pass

            for i in range(18, 67):
                try:
                    lower_dict[p_data['Project MM' + str(i)]] = p_data['Project MM' + str(i) + ' Forecast - Actual']
                except KeyError:
                    pass
        except KeyError:
            lower_dict = {}

        upper_dict[name] = lower_dict

    return upper_dict

def ap_p_milestone_data_bulk(project_list, master_data):
    upper_dict = {}

    for name in project_list:
        try:
            p_data = master_data[name]
            lower_dict = {}
            for i in range(1, 50):
                try:
                    try:
                        lower_dict[p_data['Approval MM' + str(i)]] = \
                            {p_data['Approval MM' + str(i) + ' Forecast / Actual'] : p_data['Approval MM' + str(i) + ' Notes']}
                    except KeyError:
                        lower_dict[p_data['Approval MM' + str(i)]] = \
                            {p_data['Approval MM' + str(i) + ' Forecast - Actual'] : p_data['Approval MM' + str(i) + ' Notes']}

                except KeyError:
                    pass

            for i in range(18, 67):
                try:
                    lower_dict[p_data['Project MM' + str(i)]] = \
                        {p_data['Project MM' + str(i) + ' Forecast - Actual'] : p_data['Project MM' + str(i) + ' Notes']}
                except KeyError:
                    pass
        except KeyError:
            lower_dict = {}

        upper_dict[name] = lower_dict

    return upper_dict

def project_time_difference(proj_m_data_1, proj_m_data_2):
    upper_dict = {}

    for proj_name in proj_m_data_1:
        td_dict = {}
        for milestone in proj_m_data_1[proj_name]:
            if milestone is not None:
                milestone_date = tuple(proj_m_data_1[proj_name][milestone])[0]
                try:
                    if date_of_interest <= milestone_date:
                        try:
                            old_milestone_date = tuple(proj_m_data_2[proj_name][milestone])[0]
                            time_delta = (milestone_date - old_milestone_date).days  # time_delta calculated here
                            if time_delta == 0:
                                td_dict[milestone] = 0
                            else:
                                td_dict[milestone] = time_delta
                        except (KeyError, TypeError):
                            td_dict[milestone] = 'Not reported' # not reported that quarter
                except (KeyError, TypeError):
                    td_dict[milestone] = 'No date provided' # date has now been removed

        upper_dict[proj_name] = td_dict

    return upper_dict


def filter_group(dictionary, group_of_interest):
    project_list = []
    for project in dictionary:
        if dictionary[project]['DfT Group'] == group_of_interest:
            project_list.append(project)

    return project_list

def filter_gmpp(dictionary):
    project_list = []
    for project in dictionary:
        if dictionary[project]['GMPP - IPA ID Number'] is not None:
            project_list.append(project)

    return project_list

def put_into_wb_all(project_list, t_dict, td_dict, td_dict2, wb):
    ws = wb.active

    row_num = 2
    for name in project_list:
        for i, milestone in enumerate(td_dict[name].keys()):
            ws.cell(row=row_num + i, column=1).value = name
            ws.cell(row=row_num + i, column=2).value = milestone
            try:
                milestone_date = tuple(t_dict[name][milestone])[0]
                ws.cell(row=row_num + i, column=3).value = milestone_date
            except KeyError:
                ws.cell(row=row_num + i, column=3).value = 0

            try:
                ws.cell(row=row_num + i, column=4).value = td_dict[name][milestone]
            except KeyError:
                ws.cell(row=row_num + i, column=4).value = 0
            try:
                ws.cell(row=row_num + i, column=5).value = td_dict2[name][milestone]
            except KeyError:
                ws.cell(row=row_num + i, column=5).value = 0

            try:
                milestone_date = tuple(t_dict[name][milestone])[0]
                ws.cell(row=row_num + i, column=6).value = t_dict[name][milestone][milestone_date]  # provides notes
            except IndexError:
                ws.cell(row=row_num + i, column=6).value = 0

        row_num = row_num + len(td_dict[name])

    ws.cell(row=1, column=1).value = 'Project'
    ws.cell(row=1, column=2).value = 'Milestone'
    ws.cell(row=1, column=3).value = 'Date'
    ws.cell(row=1, column=4).value = '3/m change'
    ws.cell(row=1, column=5).value = '1/y change'
    ws.cell(row=1, column=6).value = 'Notes'

    return wb


def check_m_keys_in_excel(project_list, t_dict_one, t_dict_two, t_dict_three):
    wb = Workbook()
    ws = wb.active
    red_text = Font(color="00fc2525")

    row_num = 2
    for name in project_list:
        one = list(t_dict_one[name].keys())
        [x for x in one if x is not None].sort()
        two = list(t_dict_two[name].keys())
        [x for x in two if x is not None].sort()
        three = list(t_dict_three[name].keys())
        [x for x in three if x is not None].sort()

        long = longest_list(one, two, three)
        for i in range(0, len(long)):
            ws.cell(row=row_num + i, column=1).value = name
            try:
                ws.cell(row=row_num + i, column=2).value = one[i]
            except IndexError:
                pass
            try:
                ws.cell(row=row_num + i, column=3).value = two[i]
                if two[i] not in one:
                    ws.cell(row=row_num + i, column=3).font = red_text
            except IndexError:
                pass
            try:
                ws.cell(row=row_num + i, column=4).value = three[i]
                if three[i] not in one:
                    ws.cell(row=row_num + i, column=4).font = red_text
            except IndexError:
                pass

        row_num = row_num + len(long)

    return wb



def longest_list(one, two, three):
    list_list = [one, two, three]
    #print(list_list)
    a = len(one)
    b = len(two)
    c = len(three)

    out = [a,b,c]
    out.sort()
    for x in list_list:
        if out[-1] == len(x):
            return x

# def printing_three(project_list, t_dict, td_dict, td_dict2, ordered_dict, ws):
#
#
#     #style = doc.styles['Normal']
#     #font = style.font
#     #font.name = 'Arial'
#     #font.size = Pt(10)
#
#     for name in project_list:
#         print(name)
#         doc.add_paragraph()
#
#         new_para = doc.add_paragraph()
#         sorted_dict = ordered_dict[name]
#         # print(sorted_dict)
#         heading = str(name)
#         new_para.add_run(str(heading)).bold = True
#         no_rows = len(sorted_dict) + 1
#         table1 = doc.add_table(rows=no_rows, cols=4)
#         table1.cell(0, 0).text = 'Milestone'
#         table1.cell(0, 1).text = 'Current date'
#         table1.cell(0, 2).text = 'Three month change'
#         table1.cell(0, 3).text = 'One year change'
#
#         for i, milestone in enumerate(sorted_dict):
#             table1.cell(i + 1, 0).width = Cm(8)
#             '''structured this way so that milestones are numbered'''
#             table1.cell(i + 1, 0).text = str(milestone[0]) + '. ' + str(milestone[1])
#             print(milestone)
#
#         '''place  dates into the table'''
#         for i, milestone in enumerate(sorted_dict):
#             if milestone[1] in t_dict[name]:
#                 '''date'''
#                 date = t_dict[name][milestone[1]]
#                 date = datetime.datetime.strptime(date.isoformat(), '%Y-%M-%d').strftime('%d/%M/%Y')
#                 table1.cell(i + 1, 1).text = str(date)
#                 '''time difference from last quarter'''
#                 td = td_dict[name][milestone[1]]
#                 table1.cell(i + 1, 2).text = str(td)
#                 '''time difference from oldest quarter'''
#                 td_2 = td_dict2[name][milestone[1]]
#                 table1.cell(i + 1, 3).text = str(td_2)
#
#     return doc


'''Function that runs the programme. This is also where the filtering of dates happens (understand better. date filters
are currently set at the global level. This is the standard one. It takes all milestones reporting in
current quarter and prints comparisions against those. It does not check to see whether milestones
have been removed'''
def run_standard_comparator_all(proj_list, dict_1, dict_2, dict_3):
    wb = Workbook()
    #ws = wb.active

    '''gather mini-dictionaries for each quarter'''
    current_milestones_dict = ap_p_milestone_data_bulk(proj_list, dict_1)
    print(current_milestones_dict)
    last_milestones_dict = ap_p_milestone_data_bulk(proj_list, dict_2)
    oldest_milestones_dict = ap_p_milestone_data_bulk(proj_list, dict_3)


    '''calculate time current and last quarter'''
    first_diff_dict = project_time_difference(current_milestones_dict, last_milestones_dict)
    #print(first_diff_dict)
    second_diff_dict = project_time_difference(current_milestones_dict, oldest_milestones_dict)

    run = put_into_wb_all(proj_list, current_milestones_dict, first_diff_dict, second_diff_dict, wb)
    #run = check_m_keys_in_excel(proj_list, current_milestones_dict, last_milestones_dict, oldest_milestones_dict)

    return run

''' 1) specify file path to master data '''
current_Q_dict = project_data_from_master('C:\\Users\\Standalone\\Will\\masters folder\\'
                                          'core data\\master_4_2018.xlsx')
last_Q_dict = project_data_from_master('C:\\Users\\Standalone\\Will\\masters folder\\'
                                       'core data\\master_3_2018.xlsx')
yearago_Q_dict = project_data_from_master('C:\\Users\\Standalone\\Will\\masters folder\\'
                                          'core data\\master_4_2017.xlsx')

'''all projects in portfolio'''
current_Q_list = list(current_Q_dict.keys())

'''projects by group'''
# group_names = ['Rail Group', 'HSMRPG', 'International Security and Environment', 'Roads Devolution & Motoring']
#current_Q_list = filter_gmpp(current_Q_dict)
# current_Q_list = filter_group(current_Q_dict, 'HSMRPG')
# last_Q_list = filter_group(last_Q_dict, 'HSMRPG')
# new_projects_not_reporting = [x for x in current_Q_list if x not in last_Q_list]
# current_Q_list = sorted([x for x in current_Q_list if x not in new_projects_not_reporting])

'''single project'''
#current_Q_list = ['Thameslink Programme']

'''2) Specify date after which project milestones should be returned. NOTE: Python date format is (YYYY,MM,DD)'''
date_of_interest = datetime.date(2000, 9, 1)

'''3) Specify file path to output document'''
print_miles = run_standard_comparator_all(current_Q_list, current_Q_dict, last_Q_dict, yearago_Q_dict)
print_miles.save('C:\\Users\\Standalone\\Will\\test.xlsx')