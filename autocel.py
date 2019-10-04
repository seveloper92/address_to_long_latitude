import requests
from openpyxl import load_workbook
import json
from urllib.request import urlopen
#여기서 엑셀파일 가져오고
api_key = "여기에 구글 geocoder API키만 입력하세요."
address = "서울시 강남구 역삼동 동사무소"
request = requests.get("https://maps.googleapis.com/maps/api/geocode/json?address="+address+"%20110&language=ko&sensor=false&key="+api_key)
data = request.text
search_result = json.loads(data)
print(search_result)
# 반복문으로 상세주소 뺴오기.
address_list = []
for i in range(len(search_result['results'])):
    address_list.append(search_result['results'][i]['formatted_address'])

#0번은 일반주소 1번은 도로명주소.
print(address_list[0])

geolist = []
for j in range(len(search_result['results'])):
    geolist.append(search_result['results'][j]['geometry']['location']['lat'])
    geolist.append(search_result['results'][j]['geometry']['location']['lng'])

#0~1번은 기본주소의 위경도
#2~3번은 도로명주소의 위경도.
print(geolist[0])
print(geolist[1])
#여기선 엘셀파일에 넣어준다.