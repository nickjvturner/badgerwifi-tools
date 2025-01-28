import requests
from common import REPO_BASE_URL

REPO_URL = f"{REPO_BASE_URL}/commits/main"

def get_latest_commit_info():
    try:
        response = requests.get(REPO_URL)
        response.raise_for_status()
        data = response.json()
        commit_sha = data['sha']
        commit_message = data['commit']['message']
        return commit_sha, commit_message
    except Exception as e:
        print(f"Failed to fetch latest commit info: {e}")
        return None, None


def run(self):
    message_callback = self.append_message
    message_callback('Checking for updated code on GitHub...')
    latest_sha, latest_message = get_latest_commit_info()
    message_callback(f'latest commit SHA: {latest_sha}')
    message_callback(f'latest commit message: {latest_message}')
