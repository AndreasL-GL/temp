from __future__ import unicode_literals
import openpyxl
import json, os
from datetime import date
from flask import jsonify
import datetime
if __name__ == '__main__':
    from date_functions import get_first_and_last_week_of_month,get_first_day_of_first_week, get_month_from_year_week, get_week_numbers
    from date_functions import get_month_from_week, get_month_name_from_number
    import io
    from repair_excel_file import repair_excel_file
    from get_excel_file import download_excel_file
else:
    from functions.Excel.date_functions import get_first_and_last_week_of_month,get_first_day_of_first_week, get_month_from_year_week, get_week_numbers
    from functions.Excel.date_functions import get_month_from_week, get_month_name_from_number
    import io
    from functions.Excel.repair_excel_file import repair_excel_file
    from functions.Excel.get_excel_file import download_excel_file
from werkzeug.datastructures import FileStorage

def get_dictionary_from_dagbok_sheet(sheet):
    "Creates a dictionary from a dagbook excel sheet"
    dag = [x.value for x in [x[0] for x in sheet['C4:C16']][::2]]
    fastighet = [x.value for x in [x[0] for x in sheet['A4:A16']][::2]]
    trees = [x.value for x in [x[0] for x in sheet['L5:L17']][::2]]
    name = [x.value for x in [x[0] for x in sheet['A20:A26']]]
    veckodag = [x.value for x in [x[0] for x in sheet['E20:E26']]]
    start_kl = [x.value for x in [x[0] for x in sheet['F20:F26']]]
    slut_kl = [x.value for x in [x[0] for x in sheet['G20:G26']]]
    rast_tim = [x.value for x in [x[0] for x in sheet['H20:H26']]]
    arbetspost1 = [x.value for x in [x[0] for x in sheet['I20:I26']]]  #I19 sheet['I19'].value
    arbetspost2 = [x.value for x in [x[0] for x in sheet['J20:J26']]]  #J19 sheet['J19'].value
    arbetspost3 = [x.value for x in [x[0] for x in sheet['K20:K26']]]  #K19 sheet['K19'].value
    arbetspost4 = [x.value for x in [x[0] for x in sheet['L20:L26']]]  #L19 sheet['L19'].value
    sa_arbetstid = [x.value for x in [x[0] for x in sheet['M20:M26']]] #M19 sheet['M19'].value
    resa_from = [x.value for x in [x[0] for x in sheet['F29:F35']]] #F28 sheet['F28'].value
    resa_till = [x.value for x in [x[0] for x in sheet['H29:H35']]] #H28 sheet['H28'].value
    km = [x.value for x in [x[0] for x in sheet['K29:K35']]] #K28 sheet['K28'].value
    restid = [x.value for x in [x[0] for x in sheet['L29:L35']]] #L28 sheet['L28'].value
    js = [
        {
            "beskrivning":dag, 
            "fastighet":fastighet, 
            "trees":trees,
            "personalnamn":name_1,
            "Start Kl":str(start_kl_1),
            "Slut Kl":str(slut_kl_1),
            "Rast":rast_1,
            "Veckodag":veckodag_1,
            sheet['I19'].value:arbetspost1_1,
            sheet['J19'].value:arbetspost2_1,
            sheet['K19'].value:arbetspost3_1,
            sheet['L19'].value:arbetspost4_1,
            "Resa från":resa_from_1,
            "Resa till":resa_till_1,
            "km":km_1,
            "Restid":restid_1
            } 
        for dag, fastighet, trees, name_1, start_kl_1, slut_kl_1, rast_1,veckodag_1, arbetspost1_1,
        arbetspost2_1,arbetspost3_1,arbetspost4_1,resa_from_1,resa_till_1,km_1,restid_1
        in zip(fastighet,dag,trees, name,start_kl,slut_kl,rast_tim,veckodag, arbetspost1,arbetspost2,arbetspost3,
            arbetspost4, resa_from, resa_till, km, restid)]
    js2 = []
    i=1
    for item in js:
        if not item[sheet['I19'].value]: item[sheet['I19'].value] = 0
        if not item[sheet['J19'].value]: item[sheet['J19'].value] = 0
        if not item[sheet['K19'].value]: item[sheet['K19'].value] = 0
        if not item[sheet['L19'].value]: item[sheet['L19'].value] = 0
        if "time" in str(type(item["Start Kl"]).__repr__): item[sheet['F19'].value] = item[sheet['F19'].value].strftime("%H:%M")
        if "time" in str(type(item["Slut Kl"]).__repr__): item[sheet['G19'].value] = item[sheet['G19'].value].strftime("%H:%M")
        if sheet["J2"].value:
            item["Datum"] = date.fromisocalendar(sheet['J2'].value, sheet['J3'].value, i).strftime("%Y-%m-%d")
        else: item["Datum"] = date.fromisocalendar(2023, sheet['J3'].value, i).strftime("%Y-%m-%d")
        i+=1
        js2.append(item)

        


    results={
            "info":{
                "Kontrakt nr":sheet['G2'].value,
                "Bandel":sheet['G3'].value,
                "År":sheet['J2'].value,
                "Vecka":sheet['J3'].value,
                "Lag nr":sheet['L2'].value,
                "Dag":sheet['L3'].value,
                "Dagboksnamn":sheet['D1'].value
            },
            "poster":js2
        }
    results["info"]["Sammanställning"]=sum([float(item[sheet['I19'].value]) + float(item[sheet['J19'].value]) + float(item[sheet['K19'].value]) + float(item[sheet['L19'].value]) for item in results['poster']])
    results["info"]["Månad"] = get_month_from_week(results["info"]["År"], results["info"]["Vecka"])
    return results



def convert_file_to_workbook(bytefile):
    print("------------------------",type(bytefile),'---------------------------------')
    if not isinstance(bytefile, FileStorage):
        wb = openpyxl.load_workbook(bytefile)
    else:
        bytefile.save(os.path.join(os.path.dirname(__file__),'tempfile.xlsx'))
        wb = openpyxl.load_workbook(os.path.join(os.path.dirname(__file__),'tempfile.xlsx'))

    
    
    wb,filename = call_functions(wb)
    file_data = io.BytesIO()
    wb.save(file_data)
    file_data.seek(0)
    return file_data,filename
    
    
def collect_workbook(items,filename):
    filename = items["info"]["Månad"] + " - Sammanställning - Trädexperterna"+".xlsx"
    l = download_excel_file("TrdexperternaApplikationer")
    if "finns inte" not in str(l.content):
        file_data = io.BytesIO(l.content)
        wb = openpyxl.load_workbook(file_data)
    else: 
        wb = openpyxl.load_workbook(os.path.join(os.path.dirname(__file__),'template.xlsx'))
    return wb, filename

def call_functions(wb):
    sheet = wb.active
    items = get_dictionary_from_dagbok_sheet(sheet)
    wb,filename = collect_workbook(items,'template.xlsx')
    wb = enter_items_into_sheet(wb,items)
    return wb,filename
    
    
def enter_items_into_sheet(wb, items): #Erik Rask Alstor ; Putters Alstor ; Kungsbacka skog
    """Manually enters all cells into a sheet."""
    sheet = wb.active
    # Deklarera variabler
    year = items['info']["År"]
    if not year: year=2023
    week = items['info']["Vecka"]
    
    
    # SET DATE CELLS
    first_and_last_day = get_first_and_last_week_of_month(year,get_month_from_year_week(year,week))
    daterange = excel_range_to_list("F5:AN5")
    dates = get_date_range(*first_and_last_day)
    for i,cell in enumerate(daterange):
        sheet[cell] = int(dates[i].strftime('%d'))
        
        
    # SET INDEX NUMBER OF ITEM POSTS BASED ON DATE
    date_index = {date:index for date, index in zip(dates,daterange)}
    
    # SET WEEK CELLS
    week_cells = ["F4","M4","T4","AA4","AH4"]
    weeks = get_week_numbers(dates)
    for i,cell in enumerate(week_cells):
        sheet[cell] = "V " + str(weeks[i])
        
    # SET MONTH AND TITLE
    sheet["B3"] = get_month_name_from_number(get_month_from_year_week(year=year,week=week)) + " " + "2013"
    sheet["E2"] = "Fakturaunderlag " + get_month_name_from_number(get_month_from_year_week(year=year,week=week)) + " " + "2013"
    
    
    # SET BANDEL
    if 'Bandel' in items['info'].keys():
        for cell in excel_range_to_list("B4:D4"):
            if sheet[cell] == items['info']['Bandel']: break
            elif sheet[cell].value: break
            else: 
                sheet[cell] = items['info']['Bandel']
                break
    
    # SET TIMES OF POSTS
    date_column_index = [date_index[datetime.datetime.strptime(item['Datum'],'%Y-%m-%d').date()][0] for item in items['poster']]
    for i in range(len(items['poster'])):
        items['poster'][i]['column_index'] = date_column_index[i]
        
        # SET RESTID
    if 'Platschef' in items['poster'][0].keys() or 'Trädbesiktare' in items['poster'][0].keys() or 'Träd- besiktning' in items['poster'][0].keys():
        if any([x['Restid'] for x in items['poster'] if x]):
            if sheet["AA118"].value:
                sheet["AA118"] += sum([x["Restid"] for x in items['poster'] if x['Restid']])
            else:
                sheet["AA118"] =  sum([x["Restid"] for x in items['poster'] if x['Restid']])
    else:
        if any([x['Restid'] for x in items['poster'] if x['Restid']]):
            if sheet["AA124"].value:
                sheet["AA124"] += sum([x["Restid"] for x in items['poster'] if x['Restid']])
            else:
                sheet["AA124"]  = sum([x["Restid"] for x in items['poster'] if x['Restid']])
        
        
    # SET ARBORIST TIMES AND KM
    if 'Arborist' in items['poster'][0].keys():
        if not sheet['A36'].value:
            if any([x['Arborist'] for x in items['poster'] if x!='0']):
                for cell in excel_range_to_list("A27:A36"):
                    if not sheet[cell].value or sheet[cell].value ==[x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]: 
                        sheet[cell] = [x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]
                        index_number = cell[1:]

                        cell_index = [item['column_index'] + index_number for item in items['poster']]
                        for index, arboristtimmar in zip(cell_index,[x['Arborist'] for x in items['poster']]):
                            sheet[index] = arboristtimmar

                        break
        else:
            if any([x['Arborist'] for x in items['poster'] if x!='0']):
                for cell in excel_range_to_list("A38:A47"):
                    if not sheet[cell].value or [x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]: 
                        sheet[cell] = [x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]
                        index_number = cell[1:]
                        cell_index = [item['column_index'] + index_number for item in items['poster']]
                        for index, arboristtimmar in zip(cell_index,[x['Arborist'] for x in items['poster']]):
                            sheet[index] = arboristtimmar
                        break
        if any([x['km'] for x in items['poster'] if x!='0']):
                for cell in excel_range_to_list("A49:A67"):
                    if not sheet[cell].value or sheet[cell].value ==[x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]: 
                        sheet[cell] = [x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]

                        index_number = cell[1:]
                        cell_index = [item['column_index'] + index_number for item in items['poster']]
                        for index, km in zip(cell_index,[x['km'] for x in items['poster']]):
                            sheet[index] = km
                        break
                
    # SET TRÄDBESIKTARE TIME AND KM
    if 'Trädbesiktare' in items['poster'][0].keys():
        if any([x['Trädbesiktare'] for x in items['poster'] if x!='0']):
            for cell in excel_range_to_list("A10:A16"):
                if not sheet[cell].value or sheet[cell].value ==[x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]: 
                    sheet[cell] = [x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]
                    index_number = cell[1:]

                    cell_index = [item['column_index'] + index_number for item in items['poster']]
                    for index, besiktartimmar in zip(cell_index,[x['Trädbesiktare'] for x in items['poster']]):
                        sheet[index] = besiktartimmar
                    break
            
        if any([x['km'] for x in items['poster'] if x!='0']):
                for cell in excel_range_to_list("A18:A24"):
                    if not sheet[cell].value or sheet[cell].value ==[x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]: 
                        sheet[cell] = [x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]
                        
                        index_number = cell[1:]
                        cell_index = [item['column_index'] + index_number for item in items['poster']]
                        for index, km in zip(cell_index,[x['km'] for x in items['poster']]):
                            sheet[index] = km

                        break
    if 'Träd- besiktning' in items['poster'][0].keys() and any([x['Träd- besiktning'] for x in items['poster'] if x!='0']):
        for cell in excel_range_to_list("A10:A16"):
            if not sheet[cell].value or sheet[cell].value ==[x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]: 
                sheet[cell] = [x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]
                index_number = cell[1:]
  
                cell_index = [item['column_index'] + index_number for item in items['poster']]
                for index, besiktartimmar in zip(cell_index,[x['Träd- besiktning'] for x in items['poster']]):
                    sheet[index] = besiktartimmar
                break
            
        if any([x['km'] for x in items['poster'] if x!='0']):
                for cell in excel_range_to_list("A18:A24"):
                    if not sheet[cell].value or sheet[cell].value ==[x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]: 
                        sheet[cell] = [x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]
                        
                        index_number = cell[1:]
                        cell_index = [item['column_index'] + index_number for item in items['poster']]
                        for index, km in zip(cell_index,[x['km'] for x in items['poster']]):
                            sheet[index] = km

                        break
                    
                    
    # SET Platschef Cells AND BYGGMÖTEN
    if 'Byggmöten' in items['poster'][0].keys() and any([x['Byggmöten'] for x in items['poster'] if x!='0']) and 'Platschef' in items['poster'][0].keys():
        for i,item in enumerate(items['poster']):
            items['poster'][i]['Platschef'] = items['poster'][i]['Platschef'] + items['poster'][i]['Byggmöten']
            
    if 'Platschef' in items['poster'][0].keys():
        if any([x['Platschef'] for x in items['poster'] if x!='0']):

            for cell in excel_range_to_list("A7:A8"):
                if not sheet[cell].value or sheet[cell].value ==[x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]: 
                    sheet[cell] = [x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]
                    index_number = cell[1:]

                    cell_index = [item['column_index'] + index_number for item in items['poster']]
                    for index, besiktartimmar in zip(cell_index,[x['Platschef'] for x in items['poster']]):
                        sheet[index] = besiktartimmar

                    break
            
        if any([x['km'] for x in items['poster'] if x!='0']):

                for cell in excel_range_to_list("A18:A24"):
                    if not sheet[cell].value or sheet[cell].value ==[x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]: 
                        sheet[cell] = [x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]
                        index_number = cell[1:]
                        cell_index = [item['column_index'] + index_number for item in items['poster']]
                        for index, km in zip(cell_index,[x['km'] for x in items['poster']]):
                            sheet[index] = km

                        break
                    # SET MARKARBETE TIMMAR
    if 'Mark Arb' in items['poster'][0].keys():
        if any([x['Mark Arb'] for x in items['poster'] if x!='0']):

            for cell in excel_range_to_list("A69:A76"):
                if not sheet[cell].value or sheet[cell].value ==[x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]: 
                    sheet[cell] = [x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]
                    index_number = cell[1:]

                    cell_index = [item['column_index'] + index_number for item in items['poster']]
                    for index, besiktartimmar in zip(cell_index,[x['Platschef'] for x in items['poster']]):
                        sheet[index] = besiktartimmar

                    break
                        # SET SKOTARE TIMMAR
    if 'Skotare' in items['poster'][0].keys():
        if any([x['Skotare'] for x in items['poster'] if x!='0']):

            for cell in excel_range_to_list("A93:A101"):
                if not sheet[cell].value or sheet[cell].value ==[x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]: 
                    sheet[cell] = [x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]
                    index_number = cell[1:]

                    cell_index = [item['column_index'] + index_number for item in items['poster']]
                    for index, besiktartimmar in zip(cell_index,[x['Skotare'] for x in items['poster']]):
                        sheet[index] = besiktartimmar

                    break
                
    if 'Skotning' in items['poster'][0].keys():
        if any([x['Skotning'] for x in items['poster'] if x!='0']):

            for cell in excel_range_to_list("A93:A101"):
                if not sheet[cell].value or sheet[cell].value ==[x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]: 
                    sheet[cell] = [x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]
                    index_number = cell[1:]

                    cell_index = [item['column_index'] + index_number for item in items['poster']]
                    for index, besiktartimmar in zip(cell_index,[x['Skotning'] for x in items['poster']]):
                        sheet[index] = besiktartimmar

                    break                
                
    if 'Avant med förare' in items['poster'][0].keys():
        if any([x['Avant med förare'] for x in items['poster'] if x!='0']):

            for cell in excel_range_to_list("A78:A81"):
                if not sheet[cell].value or sheet[cell].value ==[x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]: 
                    sheet[cell] = [x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]
                    index_number = cell[1:]

                    cell_index = [item['column_index'] + index_number for item in items['poster']]
                    for index, besiktartimmar in zip(cell_index,[x['Avant med förare'] for x in items['poster']]):
                        sheet[index] = besiktartimmar

                    break
                
    if 'Lastbil' in items['poster'][0].keys():
        if any([x['Lastbil'] for x in items['poster'] if x!='0']):

            for cell in excel_range_to_list("A103:A107"):
                if not sheet[cell].value or sheet[cell].value ==[x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]: 
                    sheet[cell] = [x["personalnamn"] for x in items["poster"] if x["personalnamn"]][0]
                    index_number = cell[1:]

                    cell_index = [item['column_index'] + index_number for item in items['poster']]
                    for index, besiktartimmar in zip(cell_index,[x['Lastbil'] for x in items['poster']]):
                        sheet[index] = besiktartimmar
                    break
    
                
    if 'SOS ledare' in items['poster'][0].keys():
        if sheet["E125"].value: sheet["E125"] = sheet["E125"].value + sum([float(item["SOS ledare"]) for item in items['poster']])
        else: sheet["E125"] = sum([float(item["SOS ledare"]) for item in items['poster']])
    wb.save(os.path.join(os.path.dirname(__file__),"newfile.xlsx"))
    return wb
    
    
def excel_range_to_list(range_string):

    start_cell, end_cell = range_string.split(':')
    start_col, start_row = openpyxl.utils.cell.coordinate_from_string(start_cell)
    end_col, end_row = openpyxl.utils.cell.coordinate_from_string(end_cell)
    
    cell_list = []
    for row in range(start_row, end_row+1):
        for col in range(openpyxl.utils.column_index_from_string(start_col),
                         openpyxl.utils.column_index_from_string(end_col)+1):
            cell = openpyxl.utils.cell.get_column_letter(col) + str(row)
            cell_list.append(cell)
    return cell_list

def get_date_range(start_date,end_date):
    """Creates a date range. Takes in start date and end date as inputs,
    returns a list of dates"""
    delta = datetime.timedelta(days=1)           # define the time difference between dates
    all_dates = []                      
    while start_date <= end_date:
        all_dates.append(start_date)  # add current date to the list
        start_date += delta
    return all_dates
if __name__ == '__main__':
    wb = openpyxl.load_workbook(os.path.join(os.path.dirname(__file__),'Felix.xlsx'))
    call_functions(wb)