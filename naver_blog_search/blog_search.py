"""
	네이버 검색 api를 이용해서 블로그 검색 결과 중 필요한 데이터 excel에 저장 
"""

#!python
# -*- coding: UTF-8 -*-

import cgi
import cgitb
import urllib.request
import json
import pandas as pd
import openpyxl
from datetime import datetime

cgitb.enable(display=0, logdir='log')

def main():
   
	html = f"""
	<!DOCTYPE html>
	<html>
		<head>
			<title>네이버 블로그 URL 수집</title>
		</head>
		<body style="width: 480px; margin: 0 auto;">
			<h1 style="text-align: center;">네이버 블로그 URL 수집</h1>
			<form method="post" action="blog_search.py" accept-charset="utf-8">
				<h3 style="text-align: center;">검색어</h3>
				<input type="text" id="text" name="text" style="width:480px; height:50px;">
				<br/><br/>
				<button type="submit" style="width: 480px; margin: 0 auto;">엑셀 다운</button>	
		
			</form>
		</body>
	</html>
	""";

	print("Content-type:text/html;")
	print('')
	print(html);

 	# 파라미터를 취득하기 위한 함수
	# get,post 구분없이 데이터를 가져온다.
	form = cgi.FieldStorage();
	# 파라미터 text를 취득한다.
	text = form.getvalue('text');

	if(text is not None):
		create_excel(text)

def create_excel(text):
	client_id = "Wu3cv0epXOAXpHxyVN4D"
	client_secret = ""
	encText = urllib.parse.quote(text)

	display = 100
	data = []
	for start in range(1,1002,100):
		if(start == 1001): start=1000

		url = "https://openapi.naver.com/v1/search/blog?query=" + encText+"&display="+str(display)+"&start="+str(start)
		request = urllib.request.Request(url)
		request.add_header("X-Naver-Client-Id",client_id)
		request.add_header("X-Naver-Client-Secret",client_secret)
		response = urllib.request.urlopen(request)
		rescode = response.getcode()

		if(rescode==200):
				response_body = response.read()

				jsonObject = json.loads(response_body)
				jsonArray = jsonObject.get("items")

				for index,list in enumerate(jsonArray):
					if start == 1000 and index == 0: continue

					bloggerlink = list.get('bloggerlink')
					id = bloggerlink.replace("https://blog.naver.com/","")
					email = id + '@naver.com'

					if "https://blog.naver.com/" not in bloggerlink: continue

					data.append([bloggerlink,id,email])
		else:
				print("Error Code:" + rescode)

		now = datetime.now()
		day = now.strftime('%Y%m%d')
		time = now.strftime('%H%M%S')
	
		df = pd.DataFrame(data, columns=['blog', 'id', 'email'])
		df.to_excel('D:/'+text+'_'+day+'_'+time+'.xlsx',index=False)    	
    	
if __name__ == "__main__":
    main()
