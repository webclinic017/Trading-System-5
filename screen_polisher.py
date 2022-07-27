import json  #3 imports below are for google sheets api
import gspread
from oauth2client.service_account import ServiceAccountCredentials

#google sheets api linking
credentials_file = "/Users/jakezimmerman/Documents/python3credentials.json"
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
gc = gspread.authorize(credentials)
wks = gc.open("RCP90")
screener = wks.get_worksheet(8)

#get ending row number
catalyst_cells = screener.range("D2:D1000")
ending_row = 0
for i, ticker in enumerate(catalyst_cells):
	if (ticker.value == ''):
		ending_row = i
		break
ending_row = str(ending_row + 1)
print("ending row: ", ending_row)

#put cell values into lists
catalyst_cells = screener.range("D2:D" + ending_row)
trend_cells = screener.range("F2:F" + ending_row)
cloud_score_cells = screener.range("G2:G" + ending_row)
sentiment_cells = screener.range("H2:H" + ending_row)
volume_cells = screener.range("I2:I" + ending_row)
catalyst_list, trend_list, cloud_score_list, sentiment_list, volume_list = [], [], [], [], []
print("appending to lists")
#append desired rows to lists
for i, cell in enumerate(catalyst_cells):
	if ((len(cell.value) > 1 and not cell.value[1].isupper()) or (trend_cells[i].value == '' and sentiment_cells[i].value == '0')):
		print("continuing on {}".format(cell.value))
		continue
	catalyst_list.append(cell.value)
	trend_list.append(trend_cells[i].value)
	cloud_score_list.append(cloud_score_cells[i].value)
	sentiment_list.append(sentiment_cells[i].value)
	volume_list.append(volume_cells[i].value)

print("showing catalyst list")
print(catalyst_list)
print("showing trend list")
print(trend_list)
print("clearing sheet")
#clear sheet
for cell in catalyst_cells:
	cell.value = ''
for cell in trend_cells:
	cell.value = ''
for cell in cloud_score_cells:
	cell.value = ''
for cell in sentiment_cells:
	cell.value = ''
for cell in volume_cells:
	cell.value = ''

screener.update_cells(catalyst_cells)
screener.update_cells(trend_cells)
screener.update_cells(cloud_score_cells)
screener.update_cells(sentiment_cells)
screener.update_cells(volume_cells)

#get length of new lists 
print("length of catalyst list: ", len(catalyst_list))
print("length of trend list: ", len(trend_list))
new_ending_row = str(len(catalyst_list) + 1)
print("new_ending_row: ", new_ending_row)
catalyst_cells = screener.range("D2:D" + new_ending_row)
trend_cells = screener.range("F2:F" + new_ending_row)
cloud_score_cells = screener.range("G2:G" + new_ending_row)
sentiment_cells = screener.range("H2:H" + new_ending_row)
volume_cells = screener.range("I2:I" + new_ending_row)

print("updating sheet")
#update sheet
for i, cell in enumerate(catalyst_cells):
	print("updating cell with: ", catalyst_list[i], i)
	cell.value = catalyst_list[i]
for i, cell in enumerate(trend_cells):
	cell.value = trend_list[i]
for i, cell in enumerate(cloud_score_cells):
	cell.value = cloud_score_list[i]
for i, cell in enumerate(sentiment_cells):
	cell.value = sentiment_list[i]
for i, cell in enumerate(volume_cells):
	cell.value = volume_list[i]

screener.update_cells(catalyst_cells)
screener.update_cells(trend_cells)
screener.update_cells(cloud_score_cells)
screener.update_cells(sentiment_cells)
screener.update_cells(volume_cells)

