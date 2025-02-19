import subprocess
import requests

from common import nl
from common import REPO_BASE_URL

REPO_URL = f"{REPO_BASE_URL}/commits/main"


def get_latest_commit_sha():
    try:
        response = requests.get(REPO_URL)
        response.raise_for_status()
        data = response.json()
        return data['sha']
    except Exception as e:
        print(f"Failed to fetch latest commit SHA: {e}")
        return None


def get_git_commit_sha():
    try:
        # Run the git command to get the current commit SHA
        commit_sha = subprocess.check_output(['git', 'rev-parse', 'HEAD'], stderr=subprocess.STDOUT).decode().strip()
        return commit_sha
    except subprocess.CalledProcessError as e:
        print("Error getting current git commit SHA:", e)
        return None


def check_for_updates(self):
    message_callback = self.append_message
    message_callback('Checking for updated code on GitHub...')
    latest_sha = get_latest_commit_sha()
    local_commit_sha = get_git_commit_sha()

    message_callback(f'latest commit SHA: {latest_sha}')
    message_callback(f'local commit SHA: {local_commit_sha}{nl}')

    if latest_sha and latest_sha != local_commit_sha:
        message_callback(f'New code is available on GitHub')
    elif latest_sha:
        message_callback(f'No new code available')
    else:
        message_callback(f'Unable to check for updates')


def run(message_callback):
    check_for_updates(message_callback)
