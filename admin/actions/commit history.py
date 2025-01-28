import requests
from datetime import datetime

from common import REPO_BASE_URL

REPO_URL = f"{REPO_BASE_URL}/commits"


def get_commit_info():
    try:
        response = requests.get(REPO_URL)
        response.raise_for_status()
        data = response.json()
        return data
    except Exception as e:
        print(f"Failed to fetch latest commit info: {e}")
        return None, None


def run(self):
    message_callback = self.append_message
    message_callback('Checking for updated code on GitHub...')
    commits = get_commit_info()
    # pprint(commits)

    # Sort commits by date (newest first)
    if commits:
        commits.sort(key=lambda commit: datetime.fromisoformat(commit['commit']['committer']['date'].replace("Z", "+00:00")), reverse=False)

    for commit in commits:
        sha = commit['sha']
        commit_message = commit['commit']['message']
        commit_date = commit['commit']['committer']['date']
        files_changed = len(commit.get('files', []))  # 'files' might not be present for some calls

        message_callback(f"Date: {commit_date}")
        message_callback(f"{commit_message}")
        message_callback("-" * 40)
