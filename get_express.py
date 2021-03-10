import requests
import random
def get_data(number,express_name = 'yuantong'):
    url='https://www.kuaidi100.com/query'
    headers = {
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
        'Connection': 'keep-alive',
        'Content-Length': '103',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Cookie': '',
        'DNT': '1',
        'Host': 'm.kuaidi100.com',
        'Origin': 'https://m.kuaidi100.com',
        'Referer': 'https://m.kuaidi100.com/result.jsp?com=%20%20%20%20&nu=',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-origin',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36 Edg/87.0.664.66',
        'X-Requested-With': 'XMLHttpRequest'
    }
     
    re = requests.Session()
    rre = re.get("https://m.kuaidi100.com/result.jsp?com=%20%20%20%20&nu=")
    cookies = requests.utils.dict_from_cookiejar(rre.cookies)
    headers['Cookie'] = "csrftoken={0}; WWWID={1}".format(cookies['csrftoken'],cookies['WWWID'])
    data = "postid={0}&id=1&valicode=&temp=0.55233717567226{1}&type={2}&phone=&token=&platform=MWWW".format(number,str(random.randint(1,6)),express_name)

    res=requests.post(url,data=data,headers=headers)

    result=res.json()

    return result
    

if __name__ == "__main__":
    
    data_json = get_data('YT5106522924480')
    print(data_json)

