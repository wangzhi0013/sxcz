from selenium import webdriver
from time import sleep
import xlrd

def num(t):
    data = xlrd.open_workbook('123.xlsx')
    table = data.sheet_by_name(u'Sheet1')
    cell = table.cell(t,0).value  # 获取学号
    return cell

def name(t):
    data = xlrd.open_workbook('123.xlsx')
    table = data.sheet_by_name(u'Sheet1')
    cell = table.cell(t,1).value  # 获取姓名
    return cell

def d1(t):
    data= xlrd.open_workbook('123.xlsx')
    table = data.sheet_by_name(u'Sheet1')
    cell = table.cell(t,2).value #获取理想信念分数
    return cell

def d2(t):
    data= xlrd.open_workbook('123.xlsx')
    table = data.sheet_by_name(u'Sheet1')
    cell = table.cell(t,3).value #获取社会公德分数
    return cell

def d3(t):
    data= xlrd.open_workbook('123.xlsx')
    table = data.sheet_by_name(u'Sheet1')
    cell = table.cell(t,4).value #获取职业道德分数
    return cell

def d4(t):
    data= xlrd.open_workbook('123.xlsx')
    table = data.sheet_by_name(u'Sheet1')
    cell = table.cell(t,5).value #获取遵纪守法分数
    return cell

def d5(t):
    data= xlrd.open_workbook('123.xlsx')
    table = data.sheet_by_name(u'Sheet1')
    cell = table.cell(t,6).value #获取实践教育分数
    return cell



user = input("请输入账号")
password = input("请输入密码")
d = 0
driver = webdriver.Chrome()
# 打开网页
driver.get("http://211.82.32.85/h5")
# 输入账户密码
driver.find_element_by_tag_name("input").send_keys(user)
driver.find_element_by_xpath("//*[@id='app']/div/div[2]/div[1]/div[2]/div[2]/input").send_keys(password)
# 点击登录
driver.find_element_by_xpath("//*[@id='app']/div/div[2]/div[1]/div[4]/button").click()
sleep(2)
# 点击综测互评按钮
print("点击综测互评按钮")
driver.find_element_by_xpath("//*[@id='app']/div[1]/div[3]/div/li[3]/div/a/span").click()
sleep(2)



for i in range(1,60):
    sleep(3)

#翻页
    js = "window.scrollTo(0,1800);"
    if i >= 41:
        print("翻页")
        driver.execute_script(js)
        sleep(3)
        print("翻页")
        driver.execute_script(js)
        sleep(3)
    elif i >= 21:
        print("翻页")
        driver.execute_script(js)
        sleep(3)

    sleep(3)
    #点击评分按钮
    driver.find_element_by_xpath("//*[@id='app']/div/div[3]/div/div[%s]/a/div[2]" % str(i+1)).click()
    sleep(2)

#判断data.xlsx中的数据是不是和登录用户重复
    if str(int(num(i))) == user :
        d = 1

    if d == 1 :
        i += 1

    print("开始给" + str(name(i)) + str(int(num(i))) + "评分")
    driver.find_element_by_xpath("//*[@id='app']/div/div[3]/div[3]/div[4]/div/div/div/input").clear()
    print("理想信念:"+str(d1(i)))
    driver.find_element_by_xpath("//*[@id='app']/div/div[3]/div[3]/div[4]/div/div/div/input").send_keys(str(d1(i)))
    driver.find_element_by_xpath("//*[@id='app']/div/div[3]/div[4]/div[4]/div/div/div/input").clear()
    print("社会公德:"+str(d2(i)))
    driver.find_element_by_xpath("//*[@id='app']/div/div[3]/div[4]/div[4]/div/div/div/input").send_keys(str(d2(i)))
    driver.find_element_by_xpath("//*[@id='app']/div/div[3]/div[5]/div[4]/div/div/div/input").clear()
    print("职业道德:"+str(d3(i)))
    driver.find_element_by_xpath("//*[@id='app']/div/div[3]/div[5]/div[4]/div/div/div/input").send_keys(str(d3(i)))
    driver.find_element_by_xpath("//*[@id='app']/div/div[3]/div[6]/div[4]/div/div/div/input").clear()
    print("遵纪守法:"+str(d4(i)))
    driver.find_element_by_xpath("//*[@id='app']/div/div[3]/div[6]/div[4]/div/div/div/input").send_keys(str(d4(i)))
    driver.find_element_by_xpath("//*[@id='app']/div/div[3]/div[7]/div[4]/div/div/div/input").clear()
    print("实践教育:"+str(d5(i)))
    driver.find_element_by_xpath("//*[@id='app']/div/div[3]/div[7]/div[4]/div/div/div/input").send_keys(str(d5(i)))
    driver.find_element_by_xpath("//*[@id='app']/div/div[4]/button").click()

    # 判断是否结束
    if int(num(i+1)) == 0:
        print("评分结束")
        driver.quit()



