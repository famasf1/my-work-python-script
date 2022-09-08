import requests
### Notify me when the script is completed to LINE.
def notifyme(confirmtext):
    """
    LINE Notify - Send text to my own line.
    parameter :
    confirmtext: str (required)
    """
    
    mytoken = 'kOcQyjPGgIAgTQ4qWjTlEJZFUj7GegzGefdDEiSsYJr'
    url = 'https://notify-api.line.me/api/notify'
    data = {
        'message' : confirmtext
    }
    options = {
        'Method' : 'POST',
        'Content-Type' : 'application/x-www-form-urlencoded',
        'Authorization' : f'Bearer {mytoken}',
    }
    response = requests.post(url=url, headers=options, data=data)
    print(response.status_code)