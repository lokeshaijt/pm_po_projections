"""
Persistent app settings (supplier email IDs, saved Cc list).

Streamlit Community Cloud wipes local files on every restart/redeploy, so
settings are kept as a JSON file on a separate branch of the app's GitHub
repo when a token is configured in secrets:

    [github]
    token = "github_pat_..."          # fine-grained, Contents: read & write
    repo = "owner/repo"
    branch = "app-data"               # optional, default app-data
    path = "settings.json"            # optional

Writing to a branch other than the deployed one does not trigger a
redeploy. Without a token, a local JSON file is used (not permanent on
Streamlit Cloud).
"""

import base64
import json
from pathlib import Path

import requests

API = "https://api.github.com"


class LocalStore:
    permanent = False

    def __init__(self, path):
        self.path = Path(path)

    def load(self):
        try:
            return json.loads(self.path.read_text(encoding="utf-8"))
        except (FileNotFoundError, json.JSONDecodeError):
            return None

    def save(self, data) -> None:
        self.path.write_text(json.dumps(data, indent=2, sort_keys=True), encoding="utf-8")


class GitHubStore:
    permanent = True

    def __init__(self, token, repo, branch="app-data", path="settings.json"):
        self.repo, self.branch, self.path = repo, branch, path
        self.headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
        }
        self._sha = None

    def _url(self, suffix=""):
        # No trailing slash: GitHub answers 404 for ".../repos/owner/repo/".
        return f"{API}/repos/{self.repo}" + (f"/{suffix}" if suffix else "")

    def _get_file(self):
        r = requests.get(self._url(f"contents/{self.path}"), headers=self.headers,
                         params={"ref": self.branch}, timeout=15)
        if r.status_code == 404:
            return None, None
        r.raise_for_status()
        body = r.json()
        return json.loads(base64.b64decode(body["content"]).decode("utf-8")), body["sha"]

    def _ensure_branch(self):
        r = requests.get(self._url(f"git/ref/heads/{self.branch}"), headers=self.headers, timeout=15)
        if r.status_code == 200:
            return
        if r.status_code != 404:
            r.raise_for_status()
        repo = requests.get(self._url(), headers=self.headers, timeout=15)
        repo.raise_for_status()
        default = repo.json()["default_branch"]
        base = requests.get(self._url(f"git/ref/heads/{default}"), headers=self.headers, timeout=15)
        base.raise_for_status()
        created = requests.post(self._url("git/refs"), headers=self.headers, timeout=15,
                                json={"ref": f"refs/heads/{self.branch}", "sha": base.json()["object"]["sha"]})
        if created.status_code != 422:  # 422: created meanwhile
            created.raise_for_status()

    def load(self):
        try:
            data, self._sha = self._get_file()
        except requests.HTTPError as e:
            if e.response is not None and e.response.status_code == 404:
                return None  # branch doesn't exist yet
            raise
        return data

    def save(self, data) -> None:
        content = base64.b64encode(json.dumps(data, indent=2, sort_keys=True).encode("utf-8")).decode()
        for attempt in range(2):
            if self._sha is None:
                self._ensure_branch()
                _, self._sha = self._get_file()
            payload = {"message": "Update app settings", "content": content, "branch": self.branch}
            if self._sha:
                payload["sha"] = self._sha
            r = requests.put(self._url(f"contents/{self.path}"), headers=self.headers, json=payload, timeout=15)
            if r.status_code in (409, 422) and attempt == 0:  # stale sha: refetch and retry once
                self._sha = None
                continue
            r.raise_for_status()
            self._sha = r.json()["content"]["sha"]
            return


def make_store(secrets, local_path):
    """GitHubStore when [github] token/repo are in secrets, else LocalStore."""
    try:
        gh = secrets.get("github") if "github" in secrets else None
    except Exception:  # no secrets.toml at all
        gh = None
    if gh and gh.get("token") and gh.get("repo"):
        return GitHubStore(gh["token"], gh["repo"], gh.get("branch", "app-data"), gh.get("path", "settings.json"))
    return LocalStore(local_path)
