import gspread
from oauth2client.service_account import ServiceAccountCredentials
from app.jira_reader import JiraAPI
from app.developers import developers_list


class GoogleSpreadsheet:
    def __init__(self, name_of_the_spreadsheet='Jira'):
        self.scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        self.credentials = ServiceAccountCredentials.from_json_keyfile_name('../session/client_secret.json', self.scope)
        self.client = gspread.authorize(self.credentials)
        self.name_of_the_spreadsheet = name_of_the_spreadsheet

        self.jira_api = JiraAPI()
        self.sprint_issues = self.jira_api.sprint_issues

        self.worksheet = self.create_new_worksheet(f'Sprint {self.jira_api.last_sprint_id}')

    def create_new_worksheet(self, title, rows=100, cols=30):
        spread = self.client.open(self.name_of_the_spreadsheet)

        worksheet = self.client.open(self.name_of_the_spreadsheet).add_worksheet(title, rows, cols)
        # worksheet = spread.get_worksheet(1)

        self.update_spread(spread, worksheet.id)

        # todo add functions for title, devs
        worksheet.update_acell('A1', title)
        worksheet.update_acell('A3', 'Rates per hour')

        dev_indexes = []
        for dev, pos in zip(developers_list, range(4, 4 + len(developers_list))):
            worksheet.update_acell(f'A{pos}', f'{dev}')
            dev_indexes.append(pos)
            worksheet.update_acell(f'B{pos}', 1)

        row_to_insert = ['Task', 'Task id', 'Assigned', 'Estimate (h)', 'Time spent (h)', 'Difference', 'Sum', 'Status']
        worksheet.insert_row(row_to_insert, 8)

        index_to_insert = 9
        time_spent_sum = '='
        rate_sum = '='
        for issue in self.sprint_issues:
            worksheet.update_acell(f'A{index_to_insert}', f'{issue["fields"]["summary"]}')
            worksheet.update_acell(f'B{index_to_insert}', f'{issue["id"]}')

            assignee = issue["fields"]["assignee"]["displayName"]
            worksheet.update_acell(f'C{index_to_insert}', f'{assignee}')

            estimated_time = self.convert_seconds_to_hours(issue["fields"]["timetracking"]["originalEstimateSeconds"])
            worksheet.update_acell(f'D{index_to_insert}', f'{estimated_time}')

            time_spent = self.convert_seconds_to_hours(issue["fields"]["timetracking"]["timeSpentSeconds"])
            worksheet.update_acell(f'E{index_to_insert}', '%8.2f' % time_spent)
            time_spent_sum += f'E{index_to_insert}+'

            worksheet.update_acell(f'F{index_to_insert}', f'=D{index_to_insert}-E{index_to_insert}')

            # todo deprive from for
            for i in range(len(dev_indexes)):
                if assignee == developers_list[i]:
                    worksheet.update_acell(f'G{index_to_insert}', f'=B{dev_indexes[i]}*E{index_to_insert}')
                    rate_sum += f'G{index_to_insert}+'
                    break

            worksheet.update_acell(f'H{index_to_insert}', f'{issue["fields"]["status"]["name"]}')
            index_to_insert += 1

        worksheet.update_acell(f'E{index_to_insert}', time_spent_sum[:-1])
        worksheet.update_acell(f'G{index_to_insert}', rate_sum[:-1])

        return worksheet

    @staticmethod
    def update_spread(spread, worksheet_id):
        spread.batch_update({
            "requests": [
                {
                    "mergeCells": {
                        "range": {
                            "sheetId": worksheet_id,
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
                            "sheetId": worksheet_id,
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
                            "sheetId": worksheet_id,
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
                            "sheetId": worksheet_id,
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


if __name__ == '__main__':
    spread_sheet = GoogleSpreadsheet()
