import requests

API_KEY = "IXyXem9kXGkpu7I3cOwU8QApaOWmCO6rqWmS4qUHKRAcwRQ91MYTcoZ1Gq57v2Lb"
PROJECT_ID = 51948

headers = {
    "Authorization": f"Bearer {API_KEY}",
    "Content-Type": "application/json"
}

url = f"https://ucdprj.backlog.com/api/v2/users/myself"

response = requests.get(url, headers=headers)

print(f"レスポンスコード: {response.status_code}")
print(f"レスポンス内容: {response.text}")
