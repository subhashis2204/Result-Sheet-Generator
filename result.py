from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import xlsxwriter

path="C:\Program Files (x86)\chromedriver.exe"
driver=webdriver.Chrome(path)
workbook=xlsxwriter.Workbook('CollegeResults.xlsx')
worksheet1=workbook.add_worksheet('CSE')
worksheet2=workbook.add_worksheet('IT')
worksheet3=workbook.add_worksheet('ECE')
worksheet4=workbook.add_worksheet('BT')
worksheet5=workbook.add_worksheet('AEIE')
worksheet6=workbook.add_worksheet('ChE')
worksheet7=workbook.add_worksheet('ME')
worksheet8=workbook.add_worksheet('CE')
start_roll=[12619001001,12619002001,12619003001,12619004001,12619005001,12619006001,12619007001,12619013001]
pages=[worksheet1,worksheet2,worksheet3,worksheet4,worksheet5,worksheet6,worksheet7,worksheet8]
max_count=[200,90,190,70,55,55,110,110]
driver.get("http://61.12.70.61:8084/resstude20.aspx")

row=2


for i in range(8):
	page=pages[i]
	start=start_roll[i]
	count=max_count[i]
	page.write(0,0,'uid')
	page.write(0,1,'name')
	page.write(0,2,'gsem1')
	page.write(0,3,'gsem2')
	page.write(0,4,'ygpa')

	for i in range(count):
		roll=driver.find_element_by_name("roll")
		roll.clear()
		roll.send_keys(str(start+i))

		select = Select(driver.find_element_by_name('sem'))
		select.select_by_index(2)

		roll.send_keys(Keys.RETURN)

		try:
			uid=str(start+i)
			name=(((WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.ID, "lblname")))).text.split(' ',1))[1]).strip()
			gsem1=((WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.ID, "lblbottom1")))).text.split())[3]
			gsem2=((WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.ID, "lblbottom2")))).text.split())[3]
			avg=(float(gsem1)+float(gsem2))/2
			ygpa=str(round(avg,2))
			page.write(row,0,uid)
			page.write(row,1,name)
			page.write(row,2,gsem1)
			page.write(row,3,gsem2)
			page.write(row,4,ygpa)
			print(uid,name,gsem1,gsem2,ygpa)
			row=row+1
		except:
			pass
		driver.back()
	row=2
workbook.close()
driver.close()