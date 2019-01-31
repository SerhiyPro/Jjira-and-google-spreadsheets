import requests
from requests.auth import HTTPBasicAuth
from settings import ALL_SPRINTS_BY_BOARD_URL, ALL_ISSUES_BY_SPRINTS_ID_URL
from settings import USERNAME, PASSWORD


class JiraAPI:

    def __init__(self):
        self.__last_sprint = self.get_the_last_sprint()

    @staticmethod
    def get_the_last_sprint():
        r = requests.get(ALL_SPRINTS_BY_BOARD_URL, auth=HTTPBasicAuth(USERNAME, PASSWORD))

        # todo inmprove error message
        if r.status_code != 200 or 'application/json' not in r.headers.get('content-type'):
            raise Exception(f'Incorrect status code, expected 200, got {r.status_code}')

        response = r.json()

        for sprint_info in reversed(response['values']):
            if 'completeDate' in sprint_info and sprint_info['state'] == 'closed':
                return sprint_info

    @staticmethod
    def get_issue_of_sprint_by_id(sprint_id):
        r = requests.get(f'{ALL_ISSUES_BY_SPRINTS_ID_URL}/{sprint_id}/issue', auth=HTTPBasicAuth(USERNAME, PASSWORD))

        if r.status_code != 200 or 'application/json' not in r.headers.get('content-type'):
            raise Exception(f'Incorrect status code, expected 200, got {r.status_code}')

        response = r.json()
        return response['issues']

    #todo change name
    @property
    def __last_sprint_id(self):
        return self.__last_sprint['id']

    @property
    def sprint_issues(self):
        return self.get_issue_of_sprint_by_id(self.__last_sprint_id)

    @property
    def last_sprint_id(self):
        return self.__last_sprint_id
