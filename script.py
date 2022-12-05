import re
import requests
import openpyxl 
import concurrent.futures
from bs4 import BeautifulSoup
from openpyxl.styles import Font
from .utils import usa_states, usa_states2


url = "https://mainevma.memberclicks.net/ui-directory-search/v2/search-directory-paged/"
filename = 'Maine Veterinary List 120222.xlsx'
cookie = 'serviceID=7193; 0012f0e1bd8c627a4a486a1336b31aa5=0mmvt872tume9hi5ddkm9kaqq0; Login=1; __cfruid=fb9a2c9acf1572637718e4d44673c38743057ac8-1670015244; __utma=143631657.853484218.1670015343.1670015343.1670015343.1; __utmc=143631657; __utmz=143631657.1670015343.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none)'
headers = {
  'cookie': cookie,
  'Content-Type': 'application/x-www-form-urlencoded',
  'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36'
}

wb = openpyxl.load_workbook(filename)
ws = wb.active


def scrape_page(page):
    global url, filename, wb, ws, cookie, headers
    payload = f'url=http%3A%2F%2Fservice-router.prod01.memberclicks.io%2Fsearch-results%2Fv2%2Fresults%2F82306538-7891-43df-8083-9fd6d0a80e37%3FpageSize%3D15%26pageNumber%3D{page}'
    response = requests.post(url, headers=headers, data=payload)

    for item in response.json()['results']:
        company = address = city = state = postcode = telephone = website = email = None
        top = item['top']
        if top:
            try:
                company = top[0]['html'].strip()
            except:
                pass

            try:
                address = top[1]['html'].strip() + ' ' + top[-1]['html'].strip()
                address = re.sub('\s+',' ', address).strip()
                if address == ',':
                    address = ''
                if address.startswith(' ,'):
                    address = address[2:].strip()
                elif address.startswith(','):
                    address = address[1:].strip()
            except:
                pass

            try:
                address1 = top[-1]['html'].strip()
                address1 = address1.split(',')[-1].strip().split(' ')
                state = address1[0].strip()
                postcode = address1[-1].strip()
                if postcode == state:
                    postcode = ''
                    
                if postcode == ',':
                    postcode = ''
                if postcode.startswith(','):
                    postcode = postcode[1:].strip()
            except:
                pass

        left = item['left']
        if left:
            try:
                telephone = left[0]['html'].split('</strong>')[-1].strip()
            except:
                pass

            try:
                email = left[1]['html'].strip()
            except:
                pass

            try:
                website = left[-1]['html'].strip()
            except:
                pass

        row = [company, address, city, state, postcode, telephone, website, email]
        ws.append(row)

    wb.save(filename)

pages = []
page = 1
while page <= 25:
    pages.append(page)
    page += 1

    
with concurrent.futures.ThreadPoolExecutor() as executor:
    executor.map(scrape_page, pages)
