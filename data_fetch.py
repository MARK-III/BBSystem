from openpyxl.workbook import Workbook
from openpyxl import load_workbook
import time


def _average7_weight(i, column):
    while True:
        if str(column[i].value) != 'None':
	    i += 1
	    break
	i -= 1
    t = 0
    m = 0
    for i in range(i-7, i):
        if str(column[i].value) != 'None':
	    t += column[i].value
	    m += 1
    weight = t/m
    return weight

def _find_last_n(i, column):
    j = 0
    weight_list = []
    while True:
        if str(column[i].value) != 'None':
	    weight_list.append(column[i].value)
	    j += 1
	    # how many records to get
	    if j == 6:
	        break
        i -= 1
    return list(reversed(weight_list))

def _find_current_line(column):
    current_date = time.strftime('%Y-%m-%d') + ' 00:00:00'
    i = 1
    while True:
        i += 1
	if str(column[i].value) == current_date:
	    i += 1
	    break
    return i

def get_body_weight(wb, sheet_name):
    sheet = wb[sheet_name]
    i = _find_current_line(sheet['A'])
    weight1 = _average7_weight(i, sheet['I'])
    weight2 = _average7_weight(i-7, sheet['I'])
    weight3 = _average7_weight(i-14, sheet['I'])
    return [weight3, weight2, weight1]

def _get_action_weight(wb, sheet_name, index):
    sheet = wb[sheet_name]
    i = _find_current_line(sheet['A'])
    weight_list = _find_last_n(i, sheet[index])
    return weight_list

def get_up_actions_weight(wb):
    up_actions_dict = {
                     'bench_press': 'B',
	      	     'boating': 'F',
                     'sitting_press': 'J',
		     'pull_up': 'N',
		     'triceps': 'R',
		     'biceps': 'V',
		     'curl_in': 'Z',
		     'curl_out': 'AD'
		 }   

    for key in up_actions_dict:
        up_actions_dict[key] = _get_action_weight(wb, 'UP', up_actions_dict[key])
    return up_actions_dict

def get_down_actions_weight(wb):
    down_actions_dict = {
                       'dead_lift': 'B',
		       'squatting': 'F',
		       'quadriceps': 'J',
		       'biceps_femoris': 'N',
		       'heel': 'R',
		       'abs': 'V',
		       'oblique': 'Z'
		   }
    for key in down_actions_dict:
        down_actions_dict[key] = _get_action_weight(wb, 'DOWN', down_actions_dict[key])
    return down_actions_dict

#wb = load_workbook('training.xlsx')

# To get body weight list
#body_weight_list = get_body_weight(wb, 'BMR')
#print 'body weight'
#for weight in body_weight_list:
#    print weight,
#print ''

# To get actions weight list
#up_actions_dict = get_up_actions_weight(wb)
#down_actions_dict = get_down_actions_weight(wb)
#for key in up_actions_dict:
#    print key
#    for weight in up_actions_dict[key]:
#        print weight,
#    print ''
#for key in down_actions_dict:
#    print key
#    for weight in down_actions_dict[key]:
#        print weight,
#    print ''

