from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver

import win32com.client
Excel = win32com.client.Dispatch("Excel.Application")
from selenium.webdriver.common.keys import Keys
def Exel_format(filename):
    wb = Excel.Workbooks.Open(u'D:\\gomno\\рабочий стол старая ос\\учеба\\7 семестр\\шабашка\\workscript\\'+filename+'.Xlsx')
    sheet = wb.ActiveSheet

    # получаем значение первой ячейки
    # val = sheet.Cells(1,1).value

    # получаем значения цепочки A1:A2
    vals1 = [r[0].value for r in sheet.Range("B5:B100")]
    vals2 = [r[0].value for r in sheet.Range("C5:C100")]
    vals3 = [r[0].value for r in sheet.Range("E5:E100")]
    vals4 = [r[0].value for r in sheet.Range("F5:F100")]

    # очистка поля
    for i in range(1, 10):
        for j in range(1, 100):
            sheet.Cells(j, i).value = ""

    # записываем значение таблиц в определенную ячейку
    sheet.Cells(1, 1).value = "Производитель"
    sheet.Cells(2, 1).value = "Название"
    sheet.Cells(3, 1).value = "Артикул"
    sheet.Cells(4, 1).value = "Цена"

    sheet.Cells(6, 1).value = "Насос"
    sheet.Cells(6, 13).value = "Двигатель"

    sheet.Cells(7, 1).value = "Q, m³/h"
    sheet.Cells(7, 2).value = "H, m"
    sheet.Cells(7, 3).value = "η, %"
    sheet.Cells(7, 4).value = "P₂, kW"

    sheet.Cells(7, 6).value = "H_bep"
    sheet.Cells(7, 7).value = "Q_bep"
    sheet.Cells(7, 8).value = "eta_bep"

    sheet.Cells(7, 10).value = "Q_min"
    sheet.Cells(7, 11).value = "Q_max"

    sheet.Cells(7, 13).value = "Частота вращ."
    sheet.Cells(7, 14).value = "Ном.мощность"
    sheet.Cells(7, 15).value = "Ток"
    sheet.Cells(7, 16).value = "Коэф.мощн."
    sheet.Cells(7, 17).value = "КПД 100%"
    sheet.Cells(7, 18).value = "КПД 75%"
    sheet.Cells(7, 19).value = "КПД 50%"

    # записываем последовательность
    i = 8
    for rec in vals1:
        sheet.Cells(i, 1).value = rec
        i = i + 1

    i = 8
    for rec in vals2:
        sheet.Cells(i, 2).value = rec
        i = i + 1

    i = 8
    for rec in vals3:
        sheet.Cells(i, 3).value = rec
        i = i + 1

    i = 8
    for rec in vals4:
        sheet.Cells(i, 4).value = rec
        i = i + 1
    # сохраняем рабочую книгу
    wb.Save()

    # закрываем ее
    wb.Close()

    # закрываем COM объект
    Excel.Quit()


driver = webdriver.Chrome('D:\\gomno\\рабочий стол старая ос\\учеба\\7 семестр\\шабашка\\chromedriver_win32 (1)\\chromedriver.exe')
driver.maximize_window()
wait = WebDriverWait(driver, 20)

driver.get("https://www.wilo-select.com/PumpProdConfig.aspx")
sleep(3)

driver.find_element_by_xpath(".//input[@name='loginName']").send_keys("gitvanya34")
driver.find_element_by_xpath(".//input[@name='loginPassword']").send_keys("6Cr-Sc8-Lir-inc")
driver.find_element_by_xpath(".//button[@name='ctl34']").click()


sleep(10)
wait.until(EC.element_to_be_clickable((By.XPATH,".//button[@name='ctl149']")))
driver.find_element_by_xpath(".//button[@name='ctl149']").click()
wait.until(EC.element_to_be_clickable((By.XPATH,".//input[@name='pnlBar_SelectionType$i0$i0$ddSeries']")))
driver.find_element_by_xpath(".//input[@name='pnlBar_SelectionType$i0$i0$ddSeries']").click()
driver.find_element_by_xpath(".//input[@name='pnlBar_SelectionType$i0$i0$ddSeries']").send_keys("CronoLine-IL")
driver.find_element_by_xpath(".//*[@class='rpText']").click()

sleep(10)
wait.until(EC.element_to_be_clickable((By.XPATH,".//span[@class='rmText rmExpandDown' and contains(text(),'Экспорт')]/parent::a/parent::li")))
driver.find_element_by_xpath(".//span[@class='rmText rmExpandDown' and contains(text(),'Экспорт')]/parent::a/parent::li").click()
driver.find_element_by_xpath(".//*[contains(text(),'Экспортировать точки кривой')]").click()
#редактирование файла

##
NamePump = driver.find_element_by_xpath(".//tr[ contains(@class, 'rgSelectedRow')] //div[@ class ='gridValueDivInner PRODDESC']").get_attribute("title")
ArtPump = driver.find_element_by_xpath(".//tr[ contains(@class, 'rgSelectedRow')]//div[@ class ='gridValueDivInner ARTNR']").get_attribute("title")
PricePump = driver.find_element_by_xpath(".//tr[ contains(@class, 'rgSelectedRow')]//div[@class='gridValueDivInner FITTPRICE' ]").get_attribute("title")
print(NamePump)
print(ArtPump)
print(PricePump)
##
###
infoRow_list=driver.find_elements_by_xpath(".//a[@class='mpMenuItemLbl']")#элементы списка "Информация"
infoRow_list[1].click()#выбор жлемента "Описнаие продукта"

wait.until(EC.element_to_be_clickable((By.XPATH,".//span[contains(text(),'Номинальная частота вращения')]/following-sibling::span")))#ждем появления списка

Speed_of_rotation=driver.find_element_by_xpath(".//span[contains(text(),'Номинальная частота вращения')]/following-sibling::span").text
Rated_power=driver.find_element_by_xpath(".//span[contains(text(),'Номинальная мощность')]/following-sibling::span").text
Current=driver.find_element_by_xpath(".//span[contains(text(),'Номинальный ток')]/following-sibling::span").text
Power_Factor=driver.find_element_by_xpath(".//span[contains(text(),'Коэффициент мощности')]/following-sibling::span").text
Efficiency_50=driver.find_element_by_xpath(".//span[contains(text(),'ηm 50')]/following-sibling::span").text
Efficiency_75=driver.find_element_by_xpath(".//span[contains(text(),'ηm 75')]/following-sibling::span").text
Efficiency_100= driver.find_element_by_xpath(".//span[contains(text(),'ηm 100')]/following-sibling::span").text

print(Speed_of_rotation)
print(Rated_power)
print(Current)
print(Power_Factor)
print(Efficiency_50)
print(Efficiency_75)
print(Efficiency_100)

infoRow_list=driver.find_elements_by_xpath(".//a[@class='mpMenuItemLbl']")#элементы списка "Информация"
infoRow_list[0].click()#выбор жлемента "Кривые насосов"

###
print("введите название файла")
filename= input()
Exel_format(filename)
#переход на следующую строку
pump_list = driver.find_elements_by_xpath(".//*[@class='rgRow rgEven' or @class='rgRow rgOdd']")

for pump in pump_list:

    ##Норм работающий блок
    pump.click()

    wait.until(EC.element_to_be_clickable(
        (By.XPATH, ".//span[@class='rmText rmExpandDown' and contains(text(),'Экспорт')]/parent::a/parent::li")))
    driver.find_element_by_xpath(".//span[@class='rmText rmExpandDown' and contains(text(),'Экспорт')]/parent::a/parent::li").click()

    Export_list= driver.find_elements_by_xpath(".//*[contains(text(),'Экспортировать точки кривой')]")
    Export_list[-1].click()#при выборе седующего жлемента списка образуются новые элементы

    ##

    NamePump = driver.find_element_by_xpath(
        ".//tr[ contains(@class, 'rgSelectedRow')] //div[@ class ='gridValueDivInner PRODDESC']").get_attribute("title")
    ArtPump = driver.find_element_by_xpath(
        ".//tr[ contains(@class, 'rgSelectedRow')]//div[@ class ='gridValueDivInner ARTNR']").get_attribute("title")
    PricePump = driver.find_element_by_xpath(
        ".//tr[ contains(@class, 'rgSelectedRow')]//div[@class='gridValueDivInner FITTPRICE' ]").get_attribute("title")

    print(NamePump)
    print(ArtPump)
    print(PricePump)


    infoRow_list=driver.find_elements_by_xpath(".//a[@class='mpMenuItemLbl']")#элементы списка "Информация"
    infoRow_list[1].click()#выбор жлемента "Описнаие продукта"

    sleep(5)
   # wait.until(EC.elementIsVisible((By.XPATH,".//span[contains(text(),'Номинальная частота вращения')]/following-sibling::span")))  # ждем появления списка
    Speed_of_rotation=driver.find_element_by_xpath(".//span[contains(text(),'Номинальная частота вращения')]/following-sibling::span").text

   # wait.until(EC.ElementIsVisible((By.XPATH, ".//span[contains(text(),'Номинальная мощность')]/following-sibling::span")))
    Rated_power=driver.find_element_by_xpath(".//span[contains(text(),'Номинальная мощность')]/following-sibling::span").text

    #wait.until(EC.ElementIsVisible((By.XPATH, ".//span[contains(text(),'Номинальный ток')]/following-sibling::span")))
    Current=driver.find_element_by_xpath(".//span[contains(text(),'Номинальный ток')]/following-sibling::span").text

   # wait.until(EC.ElementIsVisible((By.XPATH, ".//span[contains(text(),'Коэффициент мощности')]/following-sibling::span")))
    Power_Factor=driver.find_element_by_xpath(".//span[contains(text(),'Коэффициент мощности')]/following-sibling::span").text

   # wait.until(EC.ElementIsVisible((By.XPATH, ".//span[contains(text(),'ηm 50')]/following-sibling::span")))
    Efficiency_50=driver.find_element_by_xpath(".//span[contains(text(),'ηm 50')]/following-sibling::span").text

   # wait.until(EC.ElementIsVisible((By.XPATH, ".//span[contains(text(),'ηm 75')]/following-sibling::span")))
    Efficiency_75=driver.find_element_by_xpath(".//span[contains(text(),'ηm 75')]/following-sibling::span").text

  #  wait.until(EC.ElementIsVisible((By.XPATH, ".//span[contains(text(),'ηm 100')]/following-sibling::span")))
    Efficiency_100= driver.find_element_by_xpath(".//span[contains(text(),'ηm 100')]/following-sibling::span").text

    print(Speed_of_rotation)
    print(Rated_power)
    print(Current)
    print(Power_Factor)
    print(Efficiency_50)
    print(Efficiency_75)
    print(Efficiency_100)

    infoRow_list = driver.find_elements_by_xpath(".//a[@class='mpMenuItemLbl']")  # элементы списка "Информация"
    infoRow_list[0].click()#выбор жлемента "Кривые насосов"

    ##
    print("введите название файла")
    filename = input()
    #print(filename)

    Exel_format(filename)
    ##

##TODO 1)закинуть данные в функцию экселя,
# TODO 2)сделать ввод с консоли данных с графиков
# TODO 3) Настроить папку загрузок и папку взятия
# TODO 4) Удалить лишнии листы
# TODO 5) Переименовать файлы


##
#//div[text()=’Тема’]/parent::div
#6Cr-Sc8-Lir-inc
#gitvanya34
#.//input[@name='loginName']
#.//input[@name='loginPassword']
#driver.close()