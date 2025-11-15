import requests

api_key = "d9J1kvSFf3oFVhIJESxjJ0rKfGRkEea7Fr2K2eRPcZwU7zRzb60DOVlDanFoLfdv"
project_id = 51948  # 例: URLに出てきた project.id=51948

url_issue_types = f"https://ucdprj.backlog.com/api/v2/projects/{project_id}/issueTypes?apiKey={api_key}"
response = requests.get(url_issue_types)
issue_types = response.json()

for itype in issue_types:
    print("issueType名:", itype["name"])
    print("issueTypeId:", itype["id"])
    print("色:", itype["color"])
    print("--------")
