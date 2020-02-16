from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
import xlwt
import datetime

starttime = datetime.datetime.now()
driver = webdriver.Chrome()
driver.implicitly_wait(10) # seconds
wbk = xlwt.Workbook(encoding = 'utf-8', style_compression = 0)
wait = WebDriverWait(driver, 10)
# driver.get("https://tidesandcurrents.noaa.gov/cdata/StationList?type=Current+Data&filter=historic")
before = '//*[@id="backbutton"]'
after = '//*[@id="forwardbutton"]'
array_list = []
array_name = []
array_year = []
# **********
# file_list = open('/Users/yangnuocheng/Desktop/Currents_info/list.txt',mode='w')
# file_name = open('/Users/yangnuocheng/Desktop/Currents_info/name.txt', mode='w')
# file_year = open('/Users/yangnuocheng/Desktop/Currents_info/year.txt', mode='w')
# "num=%d" % num
# time.sleep(1)
# for i in range(1,800):
#     array_list.append(driver.find_element_by_xpath('//*[@id="stationTable"]/tbody[1]/tr[%d]/td[1]' %i).text)
#     file_list.write(array_list[-1])
#     file_list.write('\n')
#     array_name.append(driver.find_element_by_xpath('//*[@id="stationTable"]/tbody[1]/tr[%d]/td[2]' %i).text)
#     file_name.write(array_name[-1])
#     file_name.write('\n')
#     array_year.append(driver.find_element_by_xpath('//*[@id="stationTable"]/tbody[1]/tr[%d]/td[4]' %i).text)
#     file_year.write(array_year[-1])
#     file_year.write('\n')
# file_list.close()
# file_name.close()
# file_year.close()
# ***********

file_list =open('/Users/yangnuocheng/Desktop/Currents_info/list.txt',mode='r')
file_name =open('/Users/yangnuocheng/Desktop/Currents_info/name.txt',mode='r')
file_year =open('/Users/yangnuocheng/Desktop/Currents_info/year.txt',mode='r')
contents_list = file_list.readlines()
contents_name = file_name.readlines()
contents_year = file_year.readlines()
for name in contents_list:
    name = name.strip('\n')
    array_list.append(name)
for name in contents_year:
    name = name.strip('\n')
    array_year.append(name)
for name in contents_name:
    name = name.strip('\n')
    array_name.append(name)

for i in range(1,800):
    book = xlwt.Workbook(encoding='utf-8')
    name = array_name[i] + "_" + array_year[i]
    sheet = book.add_sheet("sheet1")
    basics = 0
    control = False
    driver.get("https://tidesandcurrents.noaa.gov/cdata/StationInfo?id=%s" %array_list[i])
    driver.find_element_by_xpath('//*[@id="databutton"]').click()
    # 进入了数据页面

    while(1):
        try:
            driver.find_element_by_xpath(before).click()
        except Exception as e:
            break
            # 到达数据集的最开始
    while(1):
        control=False
        control_1=False
        for i in range(0,200):
            for j in range(0,3):
                try:
                    xpath = '//*[@id="dataTable"]/tbody/tr['+str(i+1)+']/td['+str(j+1)+']'
                    ans = driver.find_element_by_xpath(xpath).text
                    print(ans)
                    sheet.write(basics+i, j, ans)
                    print("write comp")
                    # 如果这里出错，则意味着索引越界
                # except IOError:
                #     print("iioo")
                except Exception as e:
                    print(i,j)
                    control = True
                    break
                if(control):
                    basics += (i+1)
                    break
            if(control):
                control=False
                break
        try:
            driver.find_element_by_xpath(after).click()
            basics += 200
        except Exception as e:
            print(name+" is finished")
            control_1 = True
        if(control_1):
            control_1 = False
            break
    book.save('/Users/yangnuocheng/Desktop/Currents_info/%s.xls' %name)
endtime = datetime.datetime.now()
print((endtime - starttime).seconds)

print(" ****************** ")
driver.close()
