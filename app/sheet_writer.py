import gspread
from oauth2client.service_account import ServiceAccountCredentials
from time import sleep
from app.jira_reader import JiraAPI
from app.developers import developers_list


class GoogleSpreadsheet:
    def __init__(self, name_of_the_spreadsheet='Jira'):
        self.scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        self.credentials = ServiceAccountCredentials.from_json_keyfile_name('session/client_secret.json', self.scope)
        self.client = gspread.authorize(self.credentials)
        self.name_of_the_spreadsheet = name_of_the_spreadsheet
        self.insertion_index = 1

        self.jira_api = JiraAPI()
        self.sprint_issues = None
        self.spread = None
        self.worksheet = None

    def write_info_about_the_last_sprint(self):
        self.sprint_issues = self.jira_api.sprint_issues

        self.spread = self.client.open(self.name_of_the_spreadsheet)
        self.worksheet = self.create_new_worksheet(f'Sprint {self.jira_api.last_sprint_id}')
        self.fill_the_spreadsheet_in()

    def create_new_worksheet(self, title, rows=100, cols=30):
        return self.spread.add_worksheet(title, rows, cols)

    def fill_the_spreadsheet_in(self):
        self.preliminary_preparation_for_last_sprint()
        self.fill_developers_in()
        self.fill_last_sprint_info_in()

    def preliminary_preparation_for_last_sprint(self):
        self.format_spread()
        self.worksheet.update_acell(f'A{self.insertion_index}', f'Sprint {self.jira_api.last_sprint_id}')
        self.insertion_index += 2

    def fill_developers_in(self):
        self.worksheet.update_acell(f'A{self.insertion_index}', 'Rates per hour')
        self.insertion_index += 1

        for dev, pos in zip(developers_list, range(self.insertion_index, self.insertion_index + len(developers_list))):
            self.worksheet.update_acell(f'A{pos}', f'{dev}')
            self.worksheet.update_acell(f'B{pos}', 1)
            self.insertion_index += 1

    def fill_last_sprint_info_in(self):
        self.insertion_index += 1

        row_to_insert = ['Task', 'Task id', 'Assigned', 'Estimate (h)', 'Time spent (h)', 'Difference', 'Sum', 'Status']
        self.worksheet.insert_row(row_to_insert, self.insertion_index)
        self.insertion_index += 1

        time_spent_sum = '='
        rate_sum = '='
        for issue in self.sprint_issues:
            self.worksheet.update_acell(f'A{self.insertion_index}', f'{issue["fields"]["summary"]}')
            self.worksheet.update_acell(f'B{self.insertion_index}', f'{issue["id"]}')

            assignee = issue["fields"]["assignee"]["displayName"]
            self.worksheet.update_acell(f'C{self.insertion_index}', f'{assignee}')

            estimated_time = self.convert_seconds_to_hours(issue["fields"]["timetracking"]["originalEstimateSeconds"])
            self.worksheet.update_acell(f'D{self.insertion_index}', f'{estimated_time}')

            time_spent = self.convert_seconds_to_hours(issue["fields"]["timetracking"]["timeSpentSeconds"])
            self.worksheet.update_acell(f'E{self.insertion_index}', '%8.2f' % time_spent)
            time_spent_sum += f'E{self.insertion_index}+'

            self.worksheet.update_acell(f'F{self.insertion_index}', f'=D{self.insertion_index}-E{self.insertion_index}')

            try:
                assignee_cell = self.worksheet.find(assignee)
                self.worksheet.update_acell(f'G{self.insertion_index}',
                                            f'=B{assignee_cell.row}*E{self.insertion_index}')
                rate_sum += f'G{self.insertion_index}+'
            except gspread.exceptions.CellNotFound:
                self.worksheet.update_acell(f'G{self.insertion_index}', f'{assignee} is not in the list of developers')

            self.worksheet.update_acell(f'H{self.insertion_index}', f'{issue["fields"]["status"]["name"]}')
            self.insertion_index += 1

            # used in order not to exceed quota
            sleep(1)

        self.worksheet.update_acell(f'E{self.insertion_index}', time_spent_sum[:-1])
        self.worksheet.update_acell(f'G{self.insertion_index}', rate_sum[:-1])

    def format_spread(self):
        self.spread.batch_update({
            "requests": [
                {
                    "mergeCells": {
                        "range": {
                            "sheetId": self.worksheet.id,
                            "startRowIndex": 0,
                            "endRowIndex": 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": 8
                        },
                        "mergeType": "MERGE_ALL"
                    }
                },
                {
                    "repeatCell": {
                        "cell": {
                            "userEnteredFormat":
                                {
                                    "horizontalAlignment": 'CENTER',
                                    "textFormat":
                                        {
                                            "fontSize": 14
                                        }
                                }
                        },
                        "range": {
                            "sheetId": self.worksheet.id,
                            "startRowIndex": 0,
                            "endRowIndex": 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": 8
                        },
                        "fields": "userEnteredFormat"
                    }
                },
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": self.worksheet.id,
                            "dimension": "COLUMNS",
                            "startIndex": 0,
                            "endIndex": 1
                        },
                        "properties": {
                            "pixelSize": 200
                        },
                        "fields": "pixelSize"
                    }
                },
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": self.worksheet.id,
                            "dimension": "COLUMNS",
                            "startIndex": 2,
                            "endIndex": 3
                        },
                        "properties": {
                            "pixelSize": 150
                        },
                        "fields": "pixelSize"
                    }
                },
            ]})

    @staticmethod
    def convert_seconds_to_hours(seconds):
        return seconds / 3600
