from pprint import pprint
import json
import requests

url = "http://techtrade.techhead.tech/Backoffice/Branch_history/branch_history_list.aspx/Getdata"

payload = {
    "draw": 1,
    "columns": [
        {
            "data": "document_no",
            "name": "document_no",
            "searchable": True,
            "orderable": True,
            "search": {
                "value": "",
                "regex": False
            }
        },
        {
            "data": "branch_code",
            "name": "",
            "searchable": True,
            "orderable": True,
            "search": {
                "value": "",
                "regex": False
            }
        },
        {
            "data": "document_date",
            "name": "",
            "searchable": True,
            "orderable": True,
            "search": {
                "value": "",
                "regex": False
            }
        },
        {
            "data": "category_name",
            "name": "",
            "searchable": True,
            "orderable": False,
            "search": {
                "value": "",
                "regex": False
            }
        },
        {
            "data": "brand_name",
            "name": "",
            "searchable": True,
            "orderable": False,
            "search": {
                "value": "",
                "regex": False
            }
        },
        {
            "data": "series",
            "name": "",
            "searchable": True,
            "orderable": False,
            "search": {
                "value": "",
                "regex": False
            }
        },
        {
            "data": "part_number",
            "name": "",
            "searchable": True,
            "orderable": False,
            "search": {
                "value": "",
                "regex": False
            }
        },
        {
            "data": "status_name",
            "name": "status",
            "searchable": True,
            "orderable": True,
            "search": {
                "value": "",
                "regex": False
            }
        },
        {
            "data": "amount",
            "name": "",
            "searchable": True,
            "orderable": False,
            "search": {
                "value": "",
                "regex": False
            }
        },
        {
            "data": "cosmetic",
            "name": "",
            "searchable": True,
            "orderable": False,
            "search": {
                "value": "",
                "regex": False
            }
        },
        {
            "data": "gadget",
            "name": "",
            "searchable": True,
            "orderable": False,
            "search": {
                "value": "",
                "regex": False
            }
        },
        {
            "data": "voucher",
            "name": "",
            "searchable": True,
            "orderable": False,
            "search": {
                "value": "",
                "regex": False
            }
        },
        {
            "data": "invoice_no",
            "name": "",
            "searchable": True,
            "orderable": False,
            "search": {
                "value": "",
                "regex": False
            }
        },
        {
            "data": "buyer_name",
            "name": "",
            "searchable": True,
            "orderable": False,
            "search": {
                "value": "",
                "regex": False
            }
        },
        {
            "data": "campaign_name",
            "name": "",
            "searchable": True,
            "orderable": False,
            "search": {
                "value": "",
                "regex": False
            }
        },
        {
            "data": "destination_brand_name",
            "name": "",
            "searchable": True,
            "orderable": False,
            "search": {
                "value": "",
                "regex": False
            }
        },
        {
            "data": "ontop_amount",
            "name": "",
            "searchable": True,
            "orderable": False,
            "search": {
                "value": "",
                "regex": False
            }
        }
    ],
    "order": [
        {
            "column": 1,
            "dir": "desc"
        }
    ],
    "start": 0,
    "length": 1000,
    "search": {
        "value": "",
        "regex": False
    },
    "textfield": "",
    "textSearch": "",
    "textdateStart": "20/04/2022",
    "textdateEnd": "20/04/2022",
    "status": "3",
    "branchId": "0",
    "isExport": False
}
headers = {
    "User-Agent": "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:99.0) Gecko/20100101 Firefox/99.0",
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Accept-Language": "en-US,en;q=0.5",
    "Accept-Encoding": "gzip, deflate",
    "Content-Type": "application/json; charset=utf-8",
    "X-Requested-With": "XMLHttpRequest",
    "Origin": "http://techtrade.techhead.tech",
    "Connection": "keep-alive",
    "Referer": "http://techtrade.techhead.tech/Backoffice/Branch_history/branch_history_list.aspx/Getdata",
    "Cookie": "ARRAffinity=d548c3075ca9c57e8af6c1595ded6dfe6c9260f1762632d490bf3f3ac989b07a; ASP.NET_SessionId=qfaivfykylmvemaxtjruv0uy; user_name=service; usr_pwd=1234567",
    "Authorization": "Basic Og=="
}

response = requests.request("POST", url, json=payload, headers=headers)
value = json.dumps(response.json(), indent=4, separators=(',\ ', ': '))

with open('test.json', 'w+') as test:
    test.write(value)
    