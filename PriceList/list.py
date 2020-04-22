import requests
from bs4 import BeautifulSoup
import time
import re
from collections import defaultdict
import xlsxwriter

class Nifty_List:
    def __init__(self):
        self.url = 'http://farm.niftyca.com/Sell_Detail.aspx?'
        self.header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.92 Safari/537.36'}
        self.priceList = defaultdict(list)

    def getInvoice(self,ID):
        html = requests.get(self.url,params={'ifdSellID':str(ID)},headers=self.header)
        soup = BeautifulSoup(html.text,'html.parser')
        item = soup.findAll("td")
        for i in range(len(item)-1):
            if item[i].get_text() == '总计：':
                loc = i
                break
        for i in range(8,loc,8):
            #num = item[i].get_text()
            name = item[i+1].get_text()
            Spec = item[i+2].get_text()
            PerCS = item[i+3].get_text()
            #ProNo = item[i+4].get_text()
            UnitPrice = item[i+5].get_text()
            #SoldNo = item[i+6].get_text()
            #total = item[i+7].get_text()
            if name not in self.priceList:
                self.priceList[name] = [[Spec,PerCS,UnitPrice]]
            else:
                if [Spec,PerCS,UnitPrice] not in self.priceList[name]:
                    self.priceList[name].append([Spec,PerCS,UnitPrice])
            #print(name,Spec,PerCS,UnitPrice)

        return self.priceList    
        
    def Read(self):
        #test_list = [1691,1683]
        for i in range(1597,1698):
            if i != 1608: #INV1608 has server error, skip this
                print(i)
                time.sleep(5)
                self.getInvoice(i)

        if self.priceList != []:
            self.Convert2xl(self.priceList)
        else:
            print("List is empty")
    
    def Convert2xl(self,finallist):
        #print(finallist)
        workbook = xlsxwriter.Workbook('Price_List_Nifty.xlsx')
        worksheet = workbook.add_worksheet()
        row = 1
        col = 0
        worksheet.write(0,0,"Name")
        worksheet.write(0,1,"Spec")
        worksheet.write(0,2,"PerCS")
        worksheet.write(0,3,"UnitPrice")
        for key in sorted(finallist.keys()):
            worksheet.write(row,col,key)
            s_list = finallist[key]
            #print(s_list)
            inc = len(s_list)
            for i in s_list:
                pos = s_list.index(i)
                worksheet.write(row+pos,1,i[0])
                worksheet.write(row+pos,2,i[1])
                worksheet.write(row+pos,3,i[2])
            row+=inc
        
        workbook.close()

if __name__ == "__main__":
    Nlist = Nifty_List()
    Nlist.Read()