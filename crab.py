import requests,xlwt,json,re,xlsxwriter,xlwt

headers = {'Referer':'http://stockdata.stock.hexun.com/zrbg/Plate.aspx?date=2012-12-31','User-Agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36'}
urls = ['http://stockdata.stock.hexun.com/zrbg/data/zrbList.aspx?date=2012-12-31&count=20&pname=20&titType=null&page={}&callback=hxbase_json11556020538628'.format(str(i)) for i in range(1,144)]
myList = []
for i in urls:
	response = requests.get(i, headers=headers)
	response.encoding = response.apparent_encoding
	json0bj = response.text[13:-1].replace('\'','\"')
	json0bj = re.sub('<.*?>','',json0bj)
	industry = re.findall(r'industry:(.*?),', json0bj)
	industryrate = re.findall(r'industryrate:(.*?),', json0bj)
	Pricelimit = re.findall(r'Pricelimit:(.*?),', json0bj)
	stockNumber = re.findall(r'stockNumber:(.*?),', json0bj)
	lootingchips = re.findall(r'lootingchips:(.*?),', json0bj)
	Scramble = re.findall(r'Scramble:(.*?),', json0bj)
	rscramble = re.findall(r'rscramble:(.*?),', json0bj)
	Strongstock = re.findall(r'Strongstock:(.*?),', json0bj)

	for i in range(len(industry)):
		dataDict = {}
		dataDict['industry']=industry[i]
		dataDict['total']=industryrate[i]
		dataDict['rank']=Pricelimit[i]
		dataDict['holder']=stockNumber[i]
		dataDict['employee']=lootingchips[i]
		dataDict['supplier']=Scramble[i]
		dataDict['environment']=rscramble[i]
		dataDict['society']=Strongstock[i]
	
		myList.append(dataDict)


def generate_excel(expenses):
    workbook = xlsxwriter.Workbook('./rec_data.xlsx')
    worksheet = workbook.add_worksheet()
 
    # 设定格式，等号左边格式名称自定义，字典中格式为指定选项
    # bold：加粗，num_format:数字格式
    bold_format = workbook.add_format({'bold': True})
    #money_format = workbook.add_format({'num_format': '$#,##0'})
    #date_format = workbook.add_format({'num_format': 'mmmm d yyyy'})
 
    # 将二行二列设置宽度为15(从0开始)
    worksheet.set_column(1, 1, 15)
 
    # 用符号标记位置，例如：A列1行
    worksheet.write('A1', '股票和代码', bold_format)
    worksheet.write('B1', '总得分', bold_format)
    worksheet.write('C1', '等级', bold_format)
    worksheet.write('D1', '股东责任', bold_format)
    worksheet.write('E1', '员工责任', bold_format)
    worksheet.write('F1', '供应商、客户和消费者权益责任', bold_format)
    worksheet.write('G1', '环境责任', bold_format)
    worksheet.write('H1', '社会责任', bold_format)
    row = 1
    col = 0
    for item in (expenses):
            # 使用write_string方法，指定数据格式写入数据
            worksheet.write_string(row, col, item['industry'])
            worksheet.write_string(row, col + 1, item['total'])
            worksheet.write_string(row, col + 2, item['rank'])
            worksheet.write_string(row, col + 3, str(item['holder']))
            worksheet.write_string(row, col + 4, item['employee'])
            worksheet.write_string(row, col + 5, str(item['supplier']))
            worksheet.write_string(row, col + 6, item['environment'])
            worksheet.write_string(row, col + 7, item['society'])
            row += 1
    workbook.close()
 
 
if __name__ == '__main__':
    rec_data = myList
    generate_excel(rec_data)