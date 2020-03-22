import re
import urllib.request
import requests
from bs4 import BeautifulSoup
import pandas as pd
from PIL import Image
import xlsxwriter
from openpyxl.styles import Alignment
import openpyxl as op
#urllib.request.urlretrieve('https://ichef.bbci.co.uk/images/ic/1024x576/p07v6mcd.jpg','D:/Downloads/img.jpg')

import os, shutil
try:
    os.makedirs('raw images')
except:
    pass
try:
    os.makedirs('images')
except:
    pass

folder = 'raw images'
folder2 = 'images'
for the_file in os.listdir(folder):
    file_path = os.path.join(folder, the_file)
    try:
        if os.path.isfile(file_path):
            os.unlink(file_path)
        #elif os.path.isdir(file_path): shutil.rmtree(file_path)
    except Exception as e:
        print(e)
for the_file in os.listdir(folder2):
    file_path = os.path.join(folder2, the_file)
    try:
        if os.path.isfile(file_path):
            os.unlink(file_path)
        #elif os.path.isdir(file_path): shutil.rmtree(file_path)
    except Exception as e:
        print(e)
prices=[]
titles=[]
links=[]
times=[]
urllist=[]
images=[]
ind=1
baseurl=input("Write url of your craigslist search: ")
r=requests.get(baseurl)

soup=BeautifulSoup(r.content,'html.parser')

priceatag=soup.findAll('a', class_="result-image gallery")

titletag=soup.findAll(name='a', class_="result-title hdrlnk")

titletag=list(dict.fromkeys(titletag))

for ptag in priceatag:
    pricetag=ptag.find('span',class_='result-price')
    price=ptag.get_text()
    prices.append(price)

for ttag in titletag:
    title=ttag.get_text()
    titles.append(title)
    link=ttag['href']
    links.append(link)
for li in links:

    req = requests.get(li)
    soup1 = BeautifulSoup(req.content, 'html.parser')
    timetag = soup1.find('time', class_="date timeago")
    time=timetag.get_text()
    times.append(time)
    try:
        imgtag = soup1.find('a', title="1")
        imgurl = imgtag['href']
        img = urllib.request.urlretrieve(imgurl, 'raw images/img{}.jpg'.format(ind))
        images.append(img)
        ind +=1
    except:
        imgurl='http://www.4motiondarlington.org/wp-content/uploads/2013/06/No-image-found.jpg'
        img = urllib.request.urlretrieve(imgurl, 'raw images/img{}.jpg'.format(ind))
        images.append(img)
        ind += 1



lst4=[]

for i in range(len(urllist)):
    lst4.append('`')
# database---------------------------------------------------

s1 = pd.Series(links, name='Link')
s2 = pd.Series(titles, name='Title')
s3 = pd.Series(prices, name='Price')
s4 = pd.Series(times, name='Date')
s5 = pd.Series(lst4, name='Image')
df = pd.concat([s1,s2,s3,s4,s5], axis=1)


writer = pd.ExcelWriter("Output.xlsx", engine='xlsxwriter')
print(df)
ind2=1
df.to_excel(writer)
images = [x[0] for x in images]
imgs=[]
width = 300
height = 300
for image in images:
    try:
        img = Image.open(image)
        img = img.resize((width,height),Image.NEAREST)
        a1='images/photo{}.jpg'.format(ind2)
        img.save(a1)
        ind2+=1
        imgs.append(a1)
    except:
        print('no1')
        pass

workbook=writer.book


worksheet = writer.sheets['Sheet1']

worksheet.set_column('B1:B5',25)
worksheet.set_column('C1:C5',25)
worksheet.set_column('D1:D5',25)
worksheet.set_column('E1:E5',25)
worksheet.set_column('F1:F5',25)
worksheet.set_column('G1:G5',40)
worksheet.set_column('H1:H5',25)
worksheet.set_default_row(125)

img_column=5
img_row=1

for im in imgs:
    worksheet.insert_image(img_row,img_column,im,{'x_scale':0.5,'y_scale':0.5,'x_offset':5,'y_offset':5,'positioning':1})
    img_row +=1


writer.save()

wb = op.load_workbook('Output.xlsx')
worksheet = wb.active
for row_cells in worksheet.iter_rows():
    for cell in row_cells:
        cell.alignment = Alignment(horizontal='center',
                                   vertical='center')
wb.save('Output.xlsx')
