import os.path
import constants

from googleapiclient.discovery import build
from google.oauth2 import service_account

class SpreadsheetParser:
    def __init__(self, creds_file):
        self.creds_file = creds_file
        self.creds = service_account.Credentials.from_service_account_file(self.creds_file)
        self.service = build('sheets', 'v4', credentials=self.creds)

        self.sheet= self.service.spreadsheets()

    def read_values(self, range, sheet_id):
        result= self.sheet.values().get(spreadsheetId=sheet_id,
            range=range).execute()
        return result
    
    def write_values(self, range, sheet_id, body):
        result = self.service.spreadsheets().values().update(
            spreadsheetId=sheet_id, range=range,
            valueInputOption='USER_ENTERED', body=body).execute()
        print('{0} cells updated.'.format(result.get('updatedCells')))

class FinancialPlanParser(SpreadsheetParser):
    def __init__(self, creds_file, spreadsheet_id):
        super().__init__(creds_file)
        self.spreadsheet_id = spreadsheet_id
        self.etfs_map = {}

    def generate_etf_map(self):
        etfs = self.read_values(constants.ETF_SECTORS, self.spreadsheet_id)
        etfs_sectors = etfs.get('values')
        for row in etfs_sectors:
            if(len(row) > 1):
                self.etfs_map[row[0]]= row[1:]
        # print(self.etfs_map)

    def generate_position(self, portfolio_name):
        pos = self.read_values(constants.POSTIONS_RANGE[portfolio_name], self.spreadsheet_id).get('values')
        position = {}
        header = None
        for row in pos:
            if(len(row) <= 1):
                continue
            if header is None:
                header = row
            else:
                position[row[0]]= {}
                for i in range(0, len(row)):
                    position[row[0]].update({header[i] :row[i]})
        # print(position)
        return position

    def sector_portion(self, pos):
        sector_percent = {}
        for key, vals in self.etfs_map.items():
            sum_percent = 0
            for val in vals:
                sum_percent = sum_percent + float(pos.get(val, {}).get('% Of Account', '0%').strip('%'))
            sector_percent[key] = sum_percent
        print(sector_percent)
        return sector_percent

    def write_portifolio(self, portfolio_name, portfolio):
        values = [[val/100] for val in list(portfolio.values())]
        body = {
        'values': values
        }
        self.write_values(constants.PORTFOLIO_RANGE[portfolio_name], self.spreadsheet_id, body)

    def generate_portfolio(self, portfolio_name):
        pos = self.generate_position(portfolio_name)
        portfolio = self.sector_portion(pos)
        self.write_portifolio(portfolio_name, portfolio)

    def calculate(self):
        self.generate_etf_map()
        gen = self.generate_portfolio('General')
        roth_ira = self.generate_portfolio('Roth IRA')
        ira = self.generate_portfolio('IRA')

def main():
    sp = FinancialPlanParser(constants.SERVICE_ACCOUNT_FILE, constants.SPREADSHEET_ID )
    sp.calculate()


if __name__ == '__main__':
    main()