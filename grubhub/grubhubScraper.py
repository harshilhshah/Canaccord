import re, csv, time, json, requests, unicodedata

def remSpCh(s):
	s = unicodedata.normalize('NFKD', s).encode('ascii','ignore')
	return ''.join([i for i in s if ord(i) < 128])

def isRestaurantDeliveryOrPickup(delivery,pickup):
	if delivery and pickup:
		return 'Both'
	elif delivery:
		return 'Delivery Only'
	else:
		return 'Pickup Only'

def webscraper(point):
	restaurants = []
	uniq_id_arr = []
	global auth_token
	auth_url = "https://api-gtm.grubhub.com/auth"
	auth_payload = {"brand":"GRUBHUB","client_id":"beta_UmWlpstzQSFmocLy3h1UieYcVST","device_id":-1512757421,"scope":"anonymous"}
	url = "https://api-gtm.grubhub.com/restaurants/search/search_listing?orderMethod=delivery&locationMode=DELIVERY&facetSet=umami&pageSize=1000&hideHateos=true&location=POINT("+point+")&facet=variationId=default-impressionScoreBase-20160712&countOmittingTimes=true"
	headers = {'Accept':'application/json','Content-Type':'application/json','Authorization': auth_token}
	application_json = {'Content-Type':'application/json'}
	try:
		json_text = json.loads(requests.get(url,headers=headers).text)
	except:
		r = requests.post(auth_url,headers=application_json,data=json.dumps(auth_payload))
		auth_token = "Bearer " + json.loads(r.text)['session_handle']['access_token']
		headers = {'Accept':'application/json','Authorization':auth_token}
		json_text = json.loads(requests.get(url,headers=headers).text)
	if len(json_text) < 5:
		return restaurants
	for restaurant in json_text['results']:
		uniq_id = str(restaurant['restaurant_id'])
		name = remSpCh(restaurant['name'])
		addr = restaurant['address']['street_address']
		rating = str(restaurant['ratings']['rating_bayesian_half_point'])
		rating_count = str(restaurant['ratings']['rating_count'])
		min_order = '$' + str(restaurant['delivery_minimum']['price'])[:-2]
		min_delivery_fee = '$' + str(restaurant['min_delivery_fee']['price'])[:-2]
		deliveryOrPickupText = isRestaurantDeliveryOrPickup(restaurant['delivery'],restaurant['pickup'])
		if min_delivery_fee == '$':
			min_delivery_fee = 'Free'
		uniq_id_arr.append(uniq_id)
		restaurants.append([addr,name,rating,rating_count,deliveryOrPickupText,min_order,min_delivery_fee,uniq_id])
	# time to scrape restaurants from the pickup tab
	url = url.replace('delivery','pickup').replace('DELIVERY','PICKUP')
	try:
		json_text = json.loads(requests.get(url,headers=headers).text)
	except:
		r = requests.post(auth_url,headers=application_json,data=json.dumps(auth_payload))
		auth_token = "Bearer " + json.loads(r.text)['session_handle']['access_token']
		headers = {'Accept':'application/json','Authorization':auth_token}
		json_text = json.loads(requests.get(url,headers=headers).text)
	if len(json_text) < 1:
		return restaurants
	for restaurant in json_text['results']:
		uniq_id = str(restaurant['restaurant_id'])
		try:
			index = uniq_id_arr.index(uniq_id)
			restaurant[index][4] = 'Both'
		except:
			name = remSpCh(restaurant['name'])
			addr = restaurant['address']['street_address']
			rating = str(restaurant['ratings']['rating_bayesian_half_point'])
			rating_count = str(restaurant['ratings']['rating_count'])
			deliveryOrPickupText = isRestaurantDeliveryOrPickup(restaurant['delivery'],restaurant['pickup'])
			restaurants.append([addr,name,rating,rating_count,deliveryOrPickupText,'','',uniq_id])
	return restaurants

def getZipCodes(file):
	ret = []
	with open(file, 'rb') as csvfile:
		for row in csv.reader(csvfile, delimiter=',', quotechar='|'):
			ret.append(row[0:3])
	return ret

def getZipCodesWithLatLong(file):
	ret = []
	with open(file, 'rb') as csvfile:
		for row in csv.reader(csvfile, delimiter=',', quotechar='|'):
			ret.append(row[0:5])
	return ret	

def writeData(filename,format,data):
	with open(filename, format) as fp:
		a = csv.writer(fp, delimiter=',')
		a.writerow(data)

def convertZipToCoord():
	for zip_row in getZipCodes('ZipTable.csv'):
		print zip_row[0]
		loc = getLatLng(zip_row[1],zip_row[2],zip_row[0])
		if loc is None:
			continue
		big_data = [zip_row[0],zip_row[1],zip_row[2],loc[0],loc[1]]
		writeData('ZipTableWithLoc.csv',big_data)

def getLatLng(city,state,zipcode):
	if len(state) != 2:
		return None
	url = 'https://api-gtm.grubhub.com/geocode?address=' + city + ',' + state + zipcode
	headers = {'Authorization':'Bearer 207aeef9-027a-4ebf-84bc-851b3cb29274'}
	json_text = json.loads(requests.get(url,headers=headers).text)
	if len(json_text) < 1:
		return None
	lat = json_text[0]['latitude']
	lng = json_text[0]['longitude']
	return (lat,lng)

start_time = time.time()
header_row = ['Zip','City','State','Address','Restaurant Name','Rating','Number of Ratings','Delivery/Pickup','Minimum Order','Delivery Fee','Unique ID']
writeData('RestaurantData.csv','w',header_row)
auth_token = ""

for zip_with_loc_row in getZipCodesWithLatLong('ZipTableWithLoc.csv')[:]:
	zipcode = zip_with_loc_row[0]
	print zipcode
	city = zip_with_loc_row[1]
	state = zip_with_loc_row[2]
	lat = zip_with_loc_row[3]
	lng = zip_with_loc_row[4]
	point = str(lng) + '%20' + str(lat)
	try:
		for restaurant in webscraper(point):
			writeData('RestaurantData.csv','a',[zipcode,city,state] + restaurant)
	except Exception as e:
		print "*******ERROR:An error occured:*********" 
		print e
		continue

print("\n\nExecution Time: --- %2.f seconds ---\n" % (time.time() - start_time))