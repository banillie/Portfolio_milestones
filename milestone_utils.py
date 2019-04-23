
import datetime
from bcompiler.utils import project_data_from_master

'''
Function that returns all milestone names and milestone forecast/actual dates
These data sets are put into a mini dictionary, with project name at the end. 
Useful way for removing milestone data from master data - for onward manipulation

NOTE: several examples of hard coding in here, plus some work a-rounds due to master data keys not being entirely
orderly. They started that way and have caused bad habits going forward. Would be useful to tidy everything up 
at some point!
'''
def all_milestone_data_bulk(master_data):
    upper_dict = {}

    for name in master_data:
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

        upper_dict[name] = lower_dict

    return upper_dict

def all_milestone_data_single(name, master_data):
    upper_dict = {}


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

    upper_dict[name] = lower_dict

    return upper_dict

'''three quarters master dictionaries are required'''
current_Q_dict = project_data_from_master('C:\\Users\\Standalone\\Will\\masters folder\\'
                                          'core data\\master_4_2018.xlsx')
last_Q_dict = project_data_from_master('C:\\Users\\Standalone\\Will\\masters folder\\'
                                       'core data\\master_3_2018.xlsx')
yearago_Q_dict = project_data_from_master('C:\\Users\\Standalone\\Will\\masters folder\\'
                                          'core data\\master_1_2018.xlsx')

test = all_milestone_data(current_Q_dict)