import requests
import xlsxwriter
import random
import string
from datetime import datetime, timedelta
import base64

class BettingExcelList:
    def __init__(self, event_type=1):
        self.event_type = event_type
        self.base_url = base64.b64decode(b'aHR0cHM6Ly9wYi5uZXNpbmUuY29tL3YxL0JldD9ldmVudFR5cGU9MQ==').decode()

    def get_matches(self, date=None):
        if date:
            params = {'FilterDates': date}
        else:
            params = {}  

        response = requests.get(self.base_url, params=params)

        if response.status_code == 200:
            data = response.json()
            all_matches = data['d']['PopularBetList']
            
            if date:
                filtered_matches = [match for match in all_matches if match['MatchTime'].startswith(date)]
            else:
                filtered_matches = all_matches  
                
            return filtered_matches
        else:
            return None

    def write_matches_to_excel(self, matches):
        if not matches:
            print("No matches to write to Excel.")
            return

        random_suffix = ''.join(random.choices(string.ascii_letters + string.digits, k=6))
        file_name = f'nesine_matches_{random_suffix}.xlsx'
        wb = xlsxwriter.Workbook(file_name)
        ws = wb.add_worksheet('Matches')

        bold = wb.add_format({'bold': True})
        centered = wb.add_format({'align': 'center', 'valign': 'vcenter'})
        header = ['Match Code', 'Date', 'Time', 'Match Name', 'Market Name', 'Outcome Name', 'Odd', 'Played Count', 'Statistics URL']

        for col, item in enumerate(header):
            ws.write(0, col, item, bold)
            ws.set_column(col, col, len(item) + 2, centered)

        date_format = wb.add_format({'num_format': 'dd-mm-yyyy'})
        date_format.set_align('center')
        time_format = wb.add_format({'num_format': 'hh:mm:ss'})
        time_format.set_align('center')

        for row, match in enumerate(matches, start=1):
            ws.write(row, 0, match['Code'])
            match_time = datetime.strptime(match['MatchTime'], '%Y-%m-%dT%H:%M:%S')
            ws.write(row, 1, match_time.strftime('%d-%m-%Y'), date_format)
            ws.write(row, 2, match_time.strftime('%H:%M:%S'), time_format)
            ws.write(row, 3, match['Name'])
            ws.write(row, 4, match['MarketName'])
            ws.write(row, 5, match['OutcomeName'])
            ws.write(row, 6, match['Odd'])
            ws.write(row, 7, match['PlayedCount'])
            ws.write(row, 8, match['StatisticsUrl'])

        ws.set_column('A:A', 12)
        ws.set_column('B:B', 12)
        ws.set_column('C:C', 10)
        ws.set_column('D:D', 32)
        ws.set_column('E:E', 20)
        ws.set_column('F:F', 15)
        ws.set_column('G:G', 10)
        ws.set_column('H:H', 15)
        ws.set_column('I:I', 36)

        wb.close()
        print(f"Matches have been written to {file_name}")

if __name__ == "__main__":
    choice = input("Bir seçenek seçin (0: Hepsi, 1: Bugün, 2: Yarın, 3: Yarından sonraki gün) : ")

    try:
        choice = int(choice)
    except ValueError:
        print("Geçersiz bir seçenek girdiniz.")
        exit()

    if choice not in [0, 1, 2, 3]:
        print("Geçersiz bir seçenek girdiniz.")
        exit()

    selected_date = None
    if choice == 1:
        selected_date = datetime.now().strftime('%Y-%m-%d')   
    elif choice == 2:
        tomorrow = datetime.now() + timedelta(days=1)
        selected_date = tomorrow.strftime('%Y-%m-%d')
    elif choice == 3:
        day_after_tomorrow = datetime.now() + timedelta(days=2)
        selected_date = day_after_tomorrow.strftime('%Y-%m-%d')

    betting_excel_list = BettingExcelList()
    matches = betting_excel_list.get_matches(selected_date)
    
    if matches:
        betting_excel_list.write_matches_to_excel(matches)
