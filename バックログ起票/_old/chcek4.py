import requests

API_KEY = "d9J1kvSFf3oFVhIJESxjJ0rKfGRkEea7Fr2K2eRPcZwU7Rzb60DOVlDanFoLfdv"
PROJECT_ID = 51948

headers = {
    "Authorization": f"Bearer {API_KEY}",
    "Content-Type": "application/json"
}

url = f"https://ucdprj.backlog.com/api/v2/users/myself"

response = requests.get(url, headers=headers)

print(f"レスポンスコード: {response.status_code}")
print(f"レスポンス内容: {response.text}")
