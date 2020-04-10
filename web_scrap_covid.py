import requests
from bs4 import BeautifulSoup
import time
import matplotlib.pyplot as plt
import datetime
import xlwt 
from xlwt import Workbook

now = datetime.datetime.now()
print ("Data as on - " + str(now.strftime("%d/%m/%Y %H:%M:%S")))


 
URL = 'https://www.mohfw.gov.in/dashboard/index.php'
page = requests.get(URL)
soup = BeautifulSoup(page.content, 'html.parser')

#Finding active cases
print("Finding Active Cases in India...")
time.sleep(1)
active = soup.find('li', class_='bg-blue')
print("Active cases:"+ str(active.text[:5])+"\n")
time.sleep(1)

#Finding cured cases
print("Finding Cured Cases in India...")
time.sleep(1)
cured = soup.find('li', class_='bg-green')
print("Cured cases:"+ str(cured.text[:5])+"\n")
time.sleep(1)

#Findig death cases
print("Finding Death Cases in India...")
time.sleep(1)
deaths = soup.find('li', class_='bg-red')
print("Death cases:"+ str(deaths.text[:4])+"\n")
time.sleep(1)

#Finding migrated cases
print("Finding Migrated Cases in India...")
time.sleep(1)
migrated = soup.find('li', class_='bg-orange')
print("Migrated cases:"+ str(migrated.text[:3])+"\n")
time.sleep(1)

total = int(active.text[:5]) + int(cured.text[:5]) + int(deaths.text[:4]) + int(migrated.text[:3])
print("Total:"+str(total))

#Plotting Graph
fig=plt.figure()
#fig.tight_layout()
#plt.legend("Source:MoH&FW")
plt1 = fig.add_subplot(221) 
plt2 = fig.add_subplot(222) 
print("Data for last 5 mornings")
x_morning=[4067,4421,5194,5734,6412]
y_morning=[6,7,8,9,10]
plt1.plot(y_morning,x_morning,"y")
plt1.scatter(y_morning,x_morning)
#plt1.legend()
plt1.set_title("Data for last 5 mornings")
plt1.set(xlabel="Date - April Month",ylabel="Positive Cases")
#time.sleep(5)

#Plotting Graph
print("Data for last 5 evenings")
x_evening=[4281,4789,5274,5865,6761]
y_evening=[6,7,8,9,10]
plt2.plot(y_evening,x_evening,"red")
plt2.scatter(y_evening,x_evening)
#plt2.legend()
plt2.set_title("Data for last 5 evenings")
plt2.set(xlabel="Date - April Month",ylabel="Positive Cases")
plt.tight_layout()
plt.savefig('data.png')
plt.show()


  
# Workbook is created 
wb = Workbook()  
sheet1 = wb.add_sheet('COVID-19')
sheet1.write(0,0,"Date")
sheet1.write(1,0,"6 April")
sheet1.write(2,0,"7 April")
sheet1.write(3,0,"8 April")
sheet1.write(4,0,"9 April")
sheet1.write(5,0,"10 April")
sheet1.write(0,1,"Total Cases")
style = xlwt.easyxf('font: bold 1, color red;')
for i in range(len(x_evening)):
    sheet1.write(i+1,1,x_evening[i],style)
wb.save("COVID-19.xls")
print("Saving File")

