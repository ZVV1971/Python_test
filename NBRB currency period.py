# Python 3

import httplib2 as http
import json
import datetime

try:
    from urlparse import urlparse
except ImportError as err:
    from urllib.parse import urlparse

headers = {
    'Accept': 'application/json',
    'Content-Type': 'application/json; charset=UTF-8'
}

now = datetime.datetime.now()

uri = 'http://www.nbrb.by'
path = '/API/ExRates/Rates/Dynamics/298?startDate=' + str(now.year) + "-" + str(now.month - 1) + "-" + str(now.day) + '&endDate=' + str(now.year) + "-" + str(now.month) + "-" + str(now.day)

target = urlparse(uri + path)
method = 'GET'
body = ''

h = http.Http()

response, content = h.request(
        target.geturl(),
        method,
        body,
        headers)

# assume that content is a json reply
# parse content with the json module
data = json.loads(content)
print('Data', data)
print('Response' , response)
if len(data) > 0:
	print(len(data), 'parts :')
	for i in range(0, len(data)):
		json_data = json.loads(content)[i]
		print('part #' + str(i), json_data)
		print(json_data['Cur_OfficialRate'])
else:
	json_data = json.loads(content)
print(json_data['Date'])
print(len(json_data.keys()))
print(target.geturl())