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

#TODO solve problem re filtering in excle when values have + sign in front of the them
#TODO amend code so easy to run check m_keys_in_excle can be run also
#TODO create function which returns assurance milestones only

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
                        lower_dict[p_data['Approval MM' + str(i)]] = \
                            {p_data['Approval MM' + str(i) + ' Forecast / Actual']: p_data[
                                'Approval MM' + str(i) + ' Notes']}
                    except KeyError:
                        lower_dict[p_data['Approval MM' + str(i)]] = \
                            {p_data['Approval MM' + str(i) + ' Forecast - Actual']: p_data[
                                'Approval MM' + str(i) + ' Notes']}

                    lower_dict[p_data['Assurance MM' + str(i)]] = \
                        {p_data['Assurance MM' + str(i) + ' Forecast - Actual']: p_data[
                                'Assurance MM' + str(i) + ' Notes']}
                except KeyError:
                    pass

            for i in range(18, 67):
                try:
                    lower_dict[p_data['Project MM' + str(i)]] = \
                        {p_data['Project MM' + str(i) + ' Forecast - Actual']: p_data['Project MM' + str(i) + ' Notes']}
                except KeyError:
                    pass
        except KeyError:
            lower_dict = {}

        upper_dict[name] = lower_dict

    return upper_dict

'''Function to filer out approval and project delivery milestones'''
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
                value = td_dict[name][milestone]
                try:
                    if int(value) > 0:
                        ws.cell(row=row_num + i, column=4).value = '+' + str(value) + ' (days)'
                    elif int(value) < 0:
                        ws.cell(row=row_num + i, column=4).value = str(value) + ' (days)'
                    elif int(value) == 0:
                        ws.cell(row=row_num + i, column=4).value = value
                except ValueError:
                    ws.cell(row=row_num + i, column=4).value = value
            except KeyError:
                ws.cell(row=row_num + i, column=4).value = 0

            try:
                value = td_dict2[name][milestone]
                try:
                    if int(value) > 0:
                        ws.cell(row=row_num + i, column=5).value = '+' + str(value) + ' (days)'
                    elif int(value) < 0:
                        ws.cell(row=row_num + i, column=5).value = str(value) + ' (days)'
                    elif int(value) == 0:
                        ws.cell(row=row_num + i, column=5).value = value
                except ValueError:
                    ws.cell(row=row_num + i, column=5).value = value
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
    ws.cell(row=1, column=4).value = '3/m change (days)'
    ws.cell(row=1, column=5).value = '1/y change (days)'
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
    a = len(one)
    b = len(two)
    c = len(three)

    out = [a,b,c]
    out.sort()
    for x in list_list:
        if out[-1] == len(x):
            return x

'''Function that runs the programme. This is also where the filtering of dates happens (understand better. date filters
are currently set at the global level. This is the standard one. It takes all milestones reporting in
current quarter and prints comparisions against those. It does not check to see whether milestones
have been removed'''
def run_comparator(function, proj_list, dict_1, dict_2, dict_3):
    wb = Workbook()

    '''gather mini-dictionaries for each quarter'''
    current_milestones_dict = function(proj_list, dict_1)
    last_milestones_dict = function(proj_list, dict_2)
    oldest_milestones_dict = function(proj_list, dict_3)

    '''calculate time current and last quarter'''
    first_diff_dict = project_time_difference(current_milestones_dict, last_milestones_dict)
    second_diff_dict = project_time_difference(current_milestones_dict, oldest_milestones_dict)

    #run = put_into_wb_all(proj_list, current_milestones_dict, first_diff_dict, second_diff_dict, wb)
    run = check_m_keys_in_excel(proj_list, current_milestones_dict, last_milestones_dict, oldest_milestones_dict)

    return run

''' 1) specify file path to master data '''
current_Q_dict = project_data_from_master('C:\\Users\\Standalone\\Will\\masters folder\\'
                                          'core data\\Hs2_NPR_Q1_1918_draft.xlsx')
last_Q_dict = project_data_from_master('C:\\Users\\Standalone\\Will\\masters folder\\'
                                       'core data\\master_4_2018.xlsx')
yearago_Q_dict = project_data_from_master('C:\\Users\\Standalone\\Will\\masters folder\\'
                                          'core data\\master_1_2018.xlsx')

''' 2) set list of projects to be included in output. Still in development'''
'''option one - all projects'''
current_Q_list = list(current_Q_dict.keys())

'''option two - group of projects'''
# group_of_projects_list = ['Rail Group', 'HSMRPG', 'International Security and Environment', 'Roads Devolution & Motoring']

'''option three - single project'''
#one_proj_list = ['Thameslink Programme']

'''3) Specify date after which project milestones should be returned. NOTE: Python date format is (YYYY,MM,DD)'''
date_of_interest = datetime.date(2019, 1, 1)

'''4) choose type of milestones that you want to analysis through selecting appropriate function below
further instructions to follow'''
print_miles = run_comparator(all_milestone_data_bulk, current_Q_list, current_Q_dict, last_Q_dict, yearago_Q_dict)

'''5) specify file path to output document'''
print_miles.save('C:\\Users\\Standalone\\Will\\testing.xlsx')
