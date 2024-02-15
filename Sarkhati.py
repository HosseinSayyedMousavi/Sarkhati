

import threading
import requests
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QTimer, QTime
import Captcha
import Main
import Login
from selenium import webdriver
import time
import base64
import re
import os
import sys
import win32gui
from selenium.webdriver import ChromeOptions as Options
from selenium.webdriver.common.action_chains import ActionChains
from PyQt5.QtWidgets import QMessageBox
import ntplib
import math
from datetime import datetime
import webbrowser
from selenium.webdriver import FirefoxOptions as Options
try:
    os.chdir(sys._MEIPASS)
except:
    pass


def enumWindowFunc(hwnd, windowList):
    """ win32gui.EnumWindows() callback """
    global PageNameToHide
    text = win32gui.GetWindowText(hwnd)
    #className = win32gui.GetClassName(hwnd)
    if text.find(PageNameToHide) != -1 or text.find('chromedriver.exe') != -1:
        windowList.append(hwnd)

def Hider():
    global Evalue
    while True:
        if Evalue:
            quit()
        myWindows = []
        # enumerate thru all top windows and get windows which are ours
        win32gui.EnumWindows(enumWindowFunc, myWindows)

        for hwnd in myWindows:
            win32gui.ShowWindow(hwnd, False)

def Stop_hide():
        global Thread4
        global Evalue
        Evalue = True
        time.sleep(1)
        del Thread4

def Captcha_show():

    Disabler(2)
    Captcha_Form.show()
    Captcha_Ui.Captcha_label.setMovie(Captcha_Ui.movie)
    global Thread2
    Thread2 = threading.Thread(target=Prepare_Enter)
    Thread2.start()

def Prepare_Enter():
    User_Xpath='//*[@id="username"]'
    Pass_Xpath='//*[@id="password"]'
    Captcha_Xpath='//*[@id="captcha-img-plus"]'
    try:
        match Main_Ui.Kargozar.currentIndex():

            case 0:
                url = 'https://online.firouzehasia.ir/login?checkmobile=false'
            case 1:
                url = 'https://online.danayan.broker/Account/Login'
                Captcha_Xpath = '//*[@id="captcha-img"]'
            case 2 :
                url = 'https://account.emofid.com/Login'
                Captcha_Xpath='//*[@id="OLoginCaptcha_CaptchaImage"]'
                User_Xpath='//*[@id="Username"]'
                Pass_Xpath='//*[@id="Password"]'
            case 3:
                url = 'https://online.atieh-broker.ir/Login'
            case 4:
                url='https://charisma.irbrokersite.ir/'
                Captcha_Xpath='/html/body/div/div/div[2]/div[1]/div/div[1]/div/form/div[3]/div/div/div/div[1]/div/span[1]/img'
                User_Xpath='//*[@id="Username"]'
                Pass_Xpath='//*[@id="Password"]'
            case 5 :
                url='http://www.jahanonlinetrading.ir/Login'

        browser.get(url)
        try:
            WebDriverWait(browser, 10).until(EC.visibility_of_element_located(('xpath',User_Xpath)))
        except:
            Thread2 = threading.Thread(target=Prepare_Enter)
            Thread2.start()
        User_Field = browser.find_element('xpath', User_Xpath)
        global Pass_Field
        Pass_Field = browser.find_element('xpath', Pass_Xpath)
        User_Field.send_keys(Main_Ui.User_Field.text())
        Pass_Field.send_keys(Main_Ui.Password_Field.text())

        try:
            WebDriverWait(browser, 1).until(EC.visibility_of_element_located(('xpath',Captcha_Xpath)))
            img = browser.find_element('xpath', Captcha_Xpath)
            src = browser.execute_async_script("""
                var ele = arguments[0], callback = arguments[1];
                ele.addEventListener('load', function fn(){
                ele.removeEventListener('load', fn, false);
                var cnv = document.createElement('canvas');
                cnv.width = this.width+110; cnv.height = this.height+16;
                cnv.getContext('2d').drawImage(this, 0, 0);
                callback(cnv.toDataURL('image/jpeg').substring(22));
                }, false);
                ele.dispatchEvent(new Event('load'));
                """, img)
            pm = QtGui.QPixmap()
            pm.loadFromData(base64.b64decode(src))
            Captcha_Ui.Captcha_Field.setEnabled(True)
        except:
            pm = QtGui.QPixmap('Data\Images\Captcha.png')
        Captcha_Ui.Captcha_label.setPixmap(pm)
        Captcha_Ui.Enter_btn.setEnabled(True)
    except:
        Captcha_Form.close()
        Captcha_Ui.Captcha_Field.setText('')

def Enter_site():
    global Pass_Field
    Main_Ui.Sefaresh_Box.setEnabled(False)
    Main_Ui.label_14.setEnabled(False)
    Main_Ui.Vaziat_Bazaar.setText('')
    Captcha_Ui.Captcha_Field.setEnabled(False)
    Captcha_Ui.Enter_btn.setEnabled(False)
    Captcha_Ui.Captcha_label.setMovie(Captcha_Ui.movie)
    try:
        match Main_Ui.Kargozar.currentIndex():
            case 0 | 1 | 2 | 3 | 5:
                Captcha_Field_Xpath = '//*[@id="captcha"]'
                First_Win_Xpath = '//*[@id="siteVersionContainer"]//*[@class="close"]'
                Load_Xpath = '//*[@id="txt_search"]'
                Customer_Xpath = '//*[@id="customer-title-text"]'
            case 4:
                Captcha_Field_Xpath = '//*[@id="Captcha"]'
                Customer_Xpath='//*[@id="navbarDropdown"]/div/div'
                Load_Xpath = '/html/body/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div'
        try:
            Captcha_Field = browser.find_element('xpath', Captcha_Field_Xpath)
            Captcha_Field.send_keys(Captcha_Ui.Captcha_Field.text())
        except:
            pass
        Pass_Field.send_keys(Keys.ENTER)
        if Main_Ui.Kargozar.currentIndex()==2:
            WebDriverWait(browser, 5).until(EC.visibility_of_element_located(('xpath',"//div[@class='mx-auto']")))
            browser.get('https://mofidonline.com/Account/Login')
            Login =WebDriverWait(browser, 5).until(EC.visibility_of_element_located(("xpath",'//div[contains(@onclick,"openOAuthLoginPage")]')))
            Login.click()
        try:
            WebDriverWait(browser, 2).until(EC.visibility_of_element_located(('xpath',First_Win_Xpath)))
        except:
            WebDriverWait(browser, 7).until(EC.visibility_of_element_located(('xpath',Load_Xpath)))
        time.sleep(1)
        try:
            browser.find_element('xpath',First_Win_Xpath).click()
        except:
            pass

        Main_Ui.Owner_label.setText(browser.find_element('xpath',Customer_Xpath).text)


        if Main_Ui.Kargozar.currentIndex()==4:
            browser.find_element('xpath','/html/body/div/div/div[1]/div[2]/div[1]/b').click()
            time.sleep(0.5)
            browser.find_element('xpath','/html/body/div[1]/div/div[1]/div[2]/div[1]/ul/li[1]').click()
    except Exception as e:
        Captcha_Form.close()
        QMessageBox.critical(Main_Form, 'Error!', 'خطایی رخ داده است.\n لطفا مجددا تلاش کنید', QMessageBox.Ok)

    else:
        Main_Ui.Sefaresh_Box.setEnabled(True)
        Main_Ui.label_14.setEnabled(True)
        Captcha_Form.close()

def Search_Symbol():
        global Stopped_Search_Symbol
        Stopped_Search_Symbol=False
        try:
            Symbol_Search_Field = browser.find_element('xpath', '/html/body/app-container/app-content/div/div/div/div[3]/div[1]/ul[1]/li[2]/div/symbol-search/div/angucomplete/div/div[1]')
        except:
            pass
        while True:
            if Stopped_Search_Symbol:
                quit()
            time.sleep(1)
            Thistime = time.time()

            try:
                if Thistime - TIMES[-1] > 1 and Thistime - TIMES[-1] < 2 and Main_Ui.Search_Symbol.currentText() != '':
                    TIMES.append(time.time() - 10)
                    symbollist = []
                    try:
                        r = re.findall(r'.*?-', Main_Ui.Search_Symbol.currentText())[0][:-1]
                    except:
                        r = Main_Ui.Search_Symbol.currentText()


                    match Main_Ui.Kargozar.currentIndex(): # KarGozari ha
                        case 0 | 1 | 2 | 3 | 5:

                            try:
                                Symbol_Search_Field.clear()
                            except:
                                browser.find_element('xpath','/html/body/app-container/app-content/div/div/div/div[3]/div[1]/ul[1]/li[2]/div/symbol-search/div/div/span').click()
                            actions.send_keys(r)
                            actions.perform()

                            WebDriverWait(browser, 5).until(EC.visibility_of_element_located(('xpath','/html/body/app-container/app-content/div/div/div/div[3]/div[1]/ul[1]/li[2]/div/symbol-search/div/angucomplete/div/div[2]/div[3]/div')))
                            time.sleep(0.5)

                            symbols = browser.find_elements('xpath','/html/body/app-container/app-content/div/div/div/div[3]/div[1]/ul[1]/li[2]/div/symbol-search/div/angucomplete/div/div[2]/div[3]/div')
                            for symbol in symbols:
                                symbollist.append(symbol.text)
                            symbols[0].click()

                        case 4 :
                            match Main_Ui.SellOrBuy.currentIndex():
                                case 0 :
                                    if browser.current_url.find('SaleOrder')!=-1:
                                        browser.find_element('xpath','.btn-outline-dark').click()
                                case 1:
                                    if browser.current_url.find('SaleOrder')==-1:
                                        browser.find_element('xpath','.border-left-0').click()
                            WebDriverWait(browser, 15).until(EC.visibility_of_element_located(('xpath','/html/body/div[1]/div/div[2]/div[1]/div/div[2]/div/div[2]/div[2]/div/div/form/div[1]/div[1]/span[1]/span[1]/span/span[1]')))
                            time.sleep(0.5)
                            browser.find_element('xpath','/html/body/div[1]/div/div[2]/div[1]/div/div[2]/div/div[2]/div[2]/div/div/form/div[1]/div[1]/span[1]/span[1]/span/span[1]').click()
                            time.sleep(0.5)
                            Symbol_Search_Field=browser.find_element('xpath','/html/body/span/span/span[1]/input')
                            Symbol_Search_Field.send_keys(r)
                            WebDriverWait(browser, 5).until(EC.visibility_of_element_located(('xpath','/html/body/span/span/span[2]/ul/li')))
                            time.sleep(0.5)
                            symbols = browser.find_elements('xpath','/html/body/span/span/span[2]/ul/li')
                            for symbol in symbols:
                                symbollist.append(symbol.text)
                            symbols[0].click()
                    # bad az barrasi dar if :
                    Main_Ui.Search_Symbol.clear()
                    Main_Ui.Search_Symbol.addItems(symbollist)
                    getMaxMin()
                    Main_Ui.Quantity_Field.setEnabled(True)
                    Main_Ui.Price_Field.setEnabled(True)
                    Main_Ui.Total_Price.setEnabled(True)
                    Main_Ui.Setting_Box.setEnabled(True)
                    Main_Ui.SellOrBuy.setEnabled(True)
                    time.sleep(1.1)
            except:
                    pass

def Disabler(halat):
    if halat ==1 or halat ==2 or halat ==5:
        Main_Ui.Quantity_Field.setValue(0)
        Main_Ui.Price_Field.setMinimum(0)
        Main_Ui.Price_Field.setValue(0)
        Main_Ui.Setting_Box.setEnabled(False)
        Main_Ui.Price_Field.setEnabled(False)
        Main_Ui.Quantity_Field.setEnabled(False)
        Main_Ui.Max_btn.setEnabled(False)
        Main_Ui.Min_btn.setEnabled(False)
        Main_Ui.OutPut_Table.setRowCount(0)
        Main_Ui.Sarkhat_btn.setEnabled(True)
        Main_Ui.Stop_btn.setEnabled(False)
        Main_Ui.Start_Time.setTime(QTime(0,0,0,0))
        Main_Ui.End_Time.setTime(QTime(0,0,0,0))
        Main_Ui.Req_Distance.setValue(300)
    if halat==2:
        Main_Ui.Sefaresh_Box.setEnabled(False)
        Main_Ui.label_14.setEnabled(False)
        Main_Ui.Vaziat_Bazaar.setText('')
        Main_Ui.Owner_label.setText('')
        Captcha_Ui.Captcha_Field.setText('')
        Main_Ui.Search_Symbol.clearEditText()
    if halat==3:
        Main_Ui.OutPut_Table.setRowCount(0)
        Main_Ui.Sefaresh_Box.setEnabled(False)
        Main_Ui.label_14.setEnabled(False)
        Main_Ui.Vaziat_Bazaar.setText('')
        Main_Ui.Time_Label.setEnabled(False)
        Main_Ui.Vaziat_Bazaar.setEnabled(False)
        Main_Ui.Plus_btn.setEnabled(False)
        Main_Ui.Plus_btn_2.setEnabled(False)
        Main_Ui.Now_btn.setEnabled(False)
        Main_Ui.Now_btn_2.setEnabled(False)
        Main_Ui.Start_Time.setEnabled(False)
        Main_Ui.End_Time.setEnabled(False)
        Main_Ui.Req_Distance.setEnabled(False)
        Main_Ui.AutoGb.setEnabled(False)
    if halat==4:
        Main_Ui.Sefaresh_Box.setEnabled(True)
        Main_Ui.label_14.setEnabled(True)
        Main_Ui.Time_Label.setEnabled(True)
        Main_Ui.Vaziat_Bazaar.setEnabled(True)
        Main_Ui.Plus_btn.setEnabled(True)
        Main_Ui.Plus_btn_2.setEnabled(True)
        Main_Ui.Now_btn.setEnabled(True)
        Main_Ui.Now_btn_2.setEnabled(True)
        Main_Ui.Start_Time.setEnabled(True)
        Main_Ui.End_Time.setEnabled(True)
        Main_Ui.Req_Distance.setEnabled(True)
        Main_Ui.AutoGb.setEnabled(True)
    if halat==5:
        Main_Ui.Search_Symbol.clearEditText()
    if halat ==1 and Main_Ui.Search_Symbol.currentText()!='':
        Main_Ui.SellOrBuy.setEnabled(False)

def getMaxMin():
    global MaxPrice
    global MinPrice
    try:
        match Main_Ui.Kargozar.currentIndex():
            case 0|1 |2 |3 | 5 :
                browser.find_element('xpath',
                                     '/html/body/app-container/app-content/div/div/div/div[3]/div[2]/div/div/widget/div/div/div/div[2]/send-order/div/div[5]/div[2]/div/div[2]/input').click()
                WebDriverWait(browser, 15).until(
                    EC.visibility_of_element_located(('xpath',
                                                    '/html/body/app-container/app-content/div/div/div/div[3]/div[2]/div/div/widget/div/div/div/div[2]/send-order/div/div[5]/div[2]/div/div[2]/div/div[1]/span[1]'))
                )
                time.sleep(0.5)

                MaxPrice = int(browser.find_element('xpath',
                                                    '/html/body/app-container/app-content/div/div/div/div[3]/div[2]/div/div/widget/div/div/div/div[2]/send-order/div/div[5]/div[2]/div/div[2]/div/div[1]/span[1]').text.replace(
                    ',', ''))
                MinPrice = int(browser.find_element('xpath',
                                                    '/html/body/app-container/app-content/div/div/div/div[3]/div[2]/div/div/widget/div/div/div/div[2]/send-order/div/div[5]/div[2]/div/div[2]/div/div[1]/span[2]').text.replace(
                    ',', ''))
            
            case 4:
                WebDriverWait(browser, 5).until(EC.visibility_of_element_located(("css selector",'.btn-sm')))
                time.sleep(0.5)
                MaxPrice = int(browser.find_element('xpath','/html/body/div[1]/div/div[2]/div[1]/div/div[2]/div/div[2]/div[3]/div/div[2]/div[5]/span[2]/span[1]').text.replace(',', ''))
                MinPrice = int(browser.find_element('xpath','/html/body/div[1]/div/div[2]/div[1]/div/div[2]/div/div[2]/div[3]/div/div[2]/div[5]/span[2]/span[3]').text.replace(',', ''))
    except:
                MaxPrice = 0
                MinPrice = 0
                Main_Ui.Price_Field.setMaximum(9999999)
                Main_Ui.Price_Field.setMinimum(MinPrice)
                Main_Ui.Price_Field.setValue(MinPrice)
                Main_Ui.Max_btn.setEnabled(False)
                Main_Ui.Min_btn.setEnabled(False)
    else:
                Main_Ui.Max_btn.setEnabled(True)
                Main_Ui.Min_btn.setEnabled(True)
                Main_Ui.Price_Field.setMaximum(MaxPrice)
                Main_Ui.Price_Field.setMinimum(MinPrice)
                Main_Ui.Price_Field.setValue(MinPrice)

def TimeView():
    global CorrectionMSecs
    Main_Ui.Time_Label.setText(QTime.currentTime().addMSecs(CorrectionMSecs).toString('hh:mm:ss.zzz'))

def Transaction():
    global CorrectionMSecs
    global TransSuccesFul
    TransSuccesFul=False
    global StopVal


    TransTime=QTime.currentTime
    StopVal=False
    Main_Ui.Sarkhat_btn.setEnabled(False)
    Main_Ui.Stop_btn.setEnabled(True)
    match Main_Ui.Kargozar.currentIndex():
        case 0 | 1 | 2 | 3 | 5 : # Danayan va Firooze
            match Main_Ui.SellOrBuy.currentIndex():
                case 0:# Kharid
                    browser.find_element('xpath',
                                         '/html/body/app-container/app-content/div/div/div/div[3]/div[2]/div/div/widget/div/div/div/div[2]/send-order/div/div[1]/div[1]/div').click()
                case 1:#Foroosh
                    browser.find_element('xpath',
                                         '/html/body/app-container/app-content/div/div/div/div[3]/div[2]/div/div/widget/div/div/div/div[2]/send-order/div/div[1]/div[2]/div').click()
            try:
                WebDriverWait(browser, 1).until(
                        EC.visibility_of_element_located(('xpath',
                                                        '/html/body/app-container/app-content/div/div/div/div[3]/div[2]/div/div/widget/div/div/div/div[2]/send-order/div/div[7]/div[2]/a/div'))
                    )
            except:
                Transaction()
            Site_Transaction_btn = browser.find_element('xpath',
                                                        '/html/body/app-container/app-content/div/div/div/div[3]/div[2]/div/div/widget/div/div/div/div[2]/send-order/div/div[7]/div[2]/a/div')
            Site_Price_Field = browser.find_element('xpath',
                                                    '/html/body/app-container/app-content/div/div/div/div[3]/div[2]/div/div/widget/div/div/div/div[2]/send-order/div/div[5]/div[2]/div/div[2]/input')
            Site_Amount_Field = browser.find_element('xpath',
                                                     '/html/body/app-container/app-content/div/div/div/div[3]/div[2]/div/div/widget/div/div/div/div[2]/send-order/div/div[5]/div[1]/div/div[2]/input')

            try:
                r = re.findall(r'.*?-', Main_Ui.Search_Symbol.currentText())[0][:-1]
            except:
                r = Main_Ui.Search_Symbol.currentText()
            threading.Thread(target=lambda : [Site_Price_Field.clear(),Site_Price_Field.send_keys(Main_Ui.Price_Field.value())]).start()
            threading.Thread(target=lambda : [Site_Amount_Field.clear(),Site_Amount_Field.send_keys(Main_Ui.Quantity_Field.value())]).start()
            while True:
                if StopVal :
                    exit()
                Thistime1=time.time()
                if Main_Ui.Start_Time.time().msecsSinceStartOfDay() <= TransTime().addMSecs(CorrectionMSecs).msecsSinceStartOfDay() <= Main_Ui.End_Time.time().msecsSinceStartOfDay():
                    ActionTime=TransTime().addMSecs(CorrectionMSecs).toString('hh:mm:ss.zzz')
                    Site_Transaction_btn.click()
                    numrows = Main_Ui.OutPut_Table.rowCount()
                    threading.Thread(target=lambda : [Site_Price_Field.clear(),Site_Price_Field.send_keys(Main_Ui.Price_Field.value())]).start()
                    threading.Thread(target=lambda : [Site_Amount_Field.clear(),Site_Amount_Field.send_keys(Main_Ui.Quantity_Field.value())]).start()

                    match Main_Ui.AutoStopCombo.currentIndex():
                        case 0:
                            try:
                                res=browser.find_element('xpath','//*[@class="notify-item success"]').text
                                if 'موفقیت' in res:

                                    time.sleep(3)

                                    Namads=browser.find_elements('xpath','//*[@id="cell_symbol"]')

                                    Types=browser.find_elements('xpath','//*[@id="cell_orderside"]')

                                    Volumes=browser.find_elements('xpath','//*[@id="cell_qunatity"]')

                                    Prices=browser.find_elements('xpath','//*[@id="cell_orderprice"]')

                                    Times=browser.find_elements('xpath','//*[@id="cell_orderDateTime"]')

                                    Statuses=browser.find_elements('xpath', '//*[@id="cell_status"]')

                                    Turns=browser.find_elements('xpath','//*[@id="cell_OrderPlace"]')
                                    for i in range (0,len(Namads)):
                                        if Namads[i].text in r and Main_Ui.SellOrBuy.currentText() in Types[i].text and int(Prices[i].text.replace(',',''))==int(Main_Ui.Price_Field.text()) and int(Main_Ui.Quantity_Field.text())==int(Volumes[i].text.replace(',','')):
                                            Namad=Namads[i].text
                                            Type=Types[i].text
                                            Volume=Volumes[i].text
                                            Price=Prices[i].text
                                            Time=Times[i].text
                                            Turn=Turns[i].text
                                            Status=Statuses[i].text
                                            break
                                    else:
                                        raise Exception
                                else:
                                    raise Exception
                            except Exception :

                                Main_Ui.OutPut_Table.setRowCount(numrows + 1)
                                item=QtWidgets.QTableWidgetItem(str(numrows + 1))
                                item.setTextAlignment(QtCore.Qt.AlignCenter)
                                Main_Ui.OutPut_Table.setVerticalHeaderItem(numrows, item)

                                item=QtWidgets.QTableWidgetItem(QtWidgets.QTableWidgetItem(r))
                                item.setTextAlignment(QtCore.Qt.AlignCenter)
                                Main_Ui.OutPut_Table.setItem(numrows, 0, item)

                                item=QtWidgets.QTableWidgetItem(QtWidgets.QTableWidgetItem(Main_Ui.SellOrBuy.currentText()))
                                item.setTextAlignment(QtCore.Qt.AlignCenter)
                                Main_Ui.OutPut_Table.setItem(numrows, 1, item)
                                item=QtWidgets.QTableWidgetItem(QtWidgets.QTableWidgetItem(Main_Ui.Quantity_Field.text()))
                                item.setTextAlignment(QtCore.Qt.AlignCenter)
                                Main_Ui.OutPut_Table.setItem(numrows, 2, item)
                                item=QtWidgets.QTableWidgetItem(QtWidgets.QTableWidgetItem(Main_Ui.Price_Field.text()))
                                item.setTextAlignment(QtCore.Qt.AlignCenter)
                                Main_Ui.OutPut_Table.setItem(numrows, 3, item)
                                item=QtWidgets.QTableWidgetItem(QtWidgets.QTableWidgetItem(ActionTime))
                                item.setTextAlignment(QtCore.Qt.AlignCenter)
                                Main_Ui.OutPut_Table.setItem(numrows, 4,item)
                                item=QtWidgets.QTableWidgetItem(QtWidgets.QTableWidgetItem('ناموفق'))
                                item.setTextAlignment(QtCore.Qt.AlignCenter)
                                Main_Ui.OutPut_Table.setItem(numrows, 5, item)
                                item=QtWidgets.QTableWidgetItem(QtWidgets.QTableWidgetItem('---'))
                                item.setTextAlignment(QtCore.Qt.AlignCenter)
                                Main_Ui.OutPut_Table.setItem(numrows, 6, item)
                                #print(time.time()-Thistime1)
                                time.sleep(max(Main_Ui.Req_Distance.value() / 1000-(time.time()-Thistime1),0))
                                Main_Ui.OutPut_Table.scrollToBottom()
                            else:
                                Main_Ui.OutPut_Table.setRowCount(numrows + 1)
                                item=QtWidgets.QTableWidgetItem(str(numrows + 1))
                                item.setTextAlignment(QtCore.Qt.AlignCenter)
                                Main_Ui.OutPut_Table.setVerticalHeaderItem(numrows,item )

                                item=QtWidgets.QTableWidgetItem(Namad)
                                item.setTextAlignment(QtCore.Qt.AlignCenter)
                                Main_Ui.OutPut_Table.setItem(numrows, 0, item)

                                item=QtWidgets.QTableWidgetItem(Type)
                                item.setTextAlignment(QtCore.Qt.AlignCenter)
                                Main_Ui.OutPut_Table.setItem(numrows, 1, item)

                                item=QtWidgets.QTableWidgetItem(Volume)
                                item.setTextAlignment(QtCore.Qt.AlignCenter)
                                Main_Ui.OutPut_Table.setItem(numrows, 2, item)

                                item=QtWidgets.QTableWidgetItem(Price)
                                item.setTextAlignment(QtCore.Qt.AlignCenter)
                                Main_Ui.OutPut_Table.setItem(numrows, 3, item)

                                item=QtWidgets.QTableWidgetItem(Time)
                                item.setTextAlignment(QtCore.Qt.AlignCenter)
                                Main_Ui.OutPut_Table.setItem(numrows, 4, item)

                                item=QtWidgets.QTableWidgetItem(Status)
                                item.setTextAlignment(QtCore.Qt.AlignCenter)
                                Main_Ui.OutPut_Table.setItem(numrows, 5, item)

                                item=QtWidgets.QTableWidgetItem(Turn)
                                item.setTextAlignment(QtCore.Qt.AlignCenter)
                                Main_Ui.OutPut_Table.setItem(numrows, 6, item)
                                TransSuccesFul=True
                                Main_Ui.OutPut_Table.scrollToBottom()
                                Main_Ui.Stop_btn.click()
                                break
                        case 1:
                                Main_Ui.OutPut_Table.setRowCount(numrows + 1)
                                item=QtWidgets.QTableWidgetItem(str(numrows + 1))
                                item.setTextAlignment(QtCore.Qt.AlignCenter)
                                Main_Ui.OutPut_Table.setVerticalHeaderItem(numrows, item)

                                item=QtWidgets.QTableWidgetItem(QtWidgets.QTableWidgetItem(r))
                                item.setTextAlignment(QtCore.Qt.AlignCenter)
                                Main_Ui.OutPut_Table.setItem(numrows, 0, item)

                                item=QtWidgets.QTableWidgetItem(QtWidgets.QTableWidgetItem(Main_Ui.SellOrBuy.currentText()))
                                item.setTextAlignment(QtCore.Qt.AlignCenter)
                                Main_Ui.OutPut_Table.setItem(numrows, 1, item)
                                item=QtWidgets.QTableWidgetItem(QtWidgets.QTableWidgetItem(Main_Ui.Quantity_Field.text()))
                                item.setTextAlignment(QtCore.Qt.AlignCenter)
                                Main_Ui.OutPut_Table.setItem(numrows, 2, item)
                                item=QtWidgets.QTableWidgetItem(QtWidgets.QTableWidgetItem(Main_Ui.Price_Field.text()))
                                item.setTextAlignment(QtCore.Qt.AlignCenter)
                                Main_Ui.OutPut_Table.setItem(numrows, 3, item)
                                item=QtWidgets.QTableWidgetItem(QtWidgets.QTableWidgetItem(ActionTime))
                                item.setTextAlignment(QtCore.Qt.AlignCenter)
                                Main_Ui.OutPut_Table.setItem(numrows, 4,item)
                                item=QtWidgets.QTableWidgetItem(QtWidgets.QTableWidgetItem('درحال بررسی'))
                                item.setTextAlignment(QtCore.Qt.AlignCenter)
                                Main_Ui.OutPut_Table.setItem(numrows, 5, item)
                                item=QtWidgets.QTableWidgetItem(QtWidgets.QTableWidgetItem('---'))
                                item.setTextAlignment(QtCore.Qt.AlignCenter)
                                Main_Ui.OutPut_Table.setItem(numrows, 6, item)
                                #print(time.time()-Thistime1)
                                
                                time.sleep(max(Main_Ui.Req_Distance.value() / 1000-(time.time()-Thistime1),0))
                                Main_Ui.OutPut_Table.scrollToBottom()
                elif TransTime().addMSecs(CorrectionMSecs).msecsSinceStartOfDay() >= Main_Ui.End_Time.time().msecsSinceStartOfDay():
                    Main_Ui.Stop_btn.click()
                    break
        case 4:
            pass

def Starter():
    global Wily
    Wily=True
    global Thread3
    Thread3 = threading.Thread(target=Transaction)
    Thread3.start()

def TableRefresher():
    global TransSuccesFul
    numrows = Main_Ui.OutPut_Table.rowCount()

    time.sleep(3)
    try:
        Namads=browser.find_elements('xpath','//*[@id="cell_symbol"]')

        Types=browser.find_elements('xpath','//*[@id="cell_orderside"]')

        Volumes=browser.find_elements('xpath','//*[@id="cell_qunatity"]')

        Prices=browser.find_elements('xpath','//*[@id="cell_orderprice"]')

        Times=browser.find_elements('xpath','//*[@id="cell_orderDateTime"]')

        Statuses=browser.find_elements('xpath', '//*[@id="cell_status"]')

        Turns=browser.find_elements('xpath','//*[@id="cell_OrderPlace"]')
        j=0
        r = Main_Ui.Search_Symbol.currentText()
        for i in range (0,len(Namads)):
            if Namads[i].text in r and Main_Ui.SellOrBuy.currentText() in Types[i].text and int(Prices[i].text.replace(',',''))==int(Main_Ui.Price_Field.text()) and int(Main_Ui.Quantity_Field.text())==int(Volumes[i].text.replace(',','')):
                Namad=Namads[i].text
                Type=Types[i].text
                Volume=Volumes[i].text
                Price=Prices[i].text
                Time=Times[i].text
                Status=Statuses[i].text
                Turn=Turns[i].text

                item=QtWidgets.QTableWidgetItem(QtWidgets.QTableWidgetItem(Namad))
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                Main_Ui.OutPut_Table.setItem(numrows-j, 0, item)
                item=QtWidgets.QTableWidgetItem(QtWidgets.QTableWidgetItem(Type))
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                Main_Ui.OutPut_Table.setItem(numrows-j, 1, item)
                item=QtWidgets.QTableWidgetItem(QtWidgets.QTableWidgetItem(Volume))
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                Main_Ui.OutPut_Table.setItem(numrows-j, 2, item)
                item=QtWidgets.QTableWidgetItem(QtWidgets.QTableWidgetItem(Price))
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                Main_Ui.OutPut_Table.setItem(numrows-j, 3, item)
                item=QtWidgets.QTableWidgetItem(QtWidgets.QTableWidgetItem(Time))
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                Main_Ui.OutPut_Table.setItem(numrows-j, 4,item)
                item=QtWidgets.QTableWidgetItem(QtWidgets.QTableWidgetItem(Status))
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                Main_Ui.OutPut_Table.setItem(numrows-j, 5, item)
                item=QtWidgets.QTableWidgetItem(QtWidgets.QTableWidgetItem(Turn))
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                Main_Ui.OutPut_Table.setItem(numrows-j, 6, item)
                j+=1
        for z in range(j,numrows+1):
            item=QtWidgets.QTableWidgetItem(QtWidgets.QTableWidgetItem('ناموفق'))
            item.setTextAlignment(QtCore.Qt.AlignCenter)
            Main_Ui.OutPut_Table.setItem(numrows-z, 5, item)
        if j>0:
            TransSuccesFul=True
        Main_Ui.OutPut_Table.scrollToTop()
        time.sleep(0.1)
        Main_Ui.OutPut_Table.scrollToBottom()
    except:
        for z in range(1,numrows+1):
            item=QtWidgets.QTableWidgetItem(QtWidgets.QTableWidgetItem('ناموفق'))
            item.setTextAlignment(QtCore.Qt.AlignCenter)
            Main_Ui.OutPut_Table.setItem(numrows-z, 6, item)
    
        Main_Ui.OutPut_Table.scrollToTop()
        time.sleep(0.1)
        Main_Ui.OutPut_Table.scrollToBottom()

def Stopper():
    if Main_Ui.AutoStopCombo.currentIndex()==1:
        threading.Thread(target=TableRefresher).start()

    if TransSuccesFul:
        QMessageBox.information(Main_Form, 'Succesful!', 'خرید شما با موفقیت انجام شد!', QMessageBox.Ok)

    global Wily
    Wily=False
    Main_Ui.Stop_btn.setEnabled(False)
    global StopVal
    StopVal = True
    time.sleep(0.5)
    global Thread3
    try:
        del Thread3
    except:
        pass
    Main_Ui.Sarkhat_btn.setEnabled(True)

def GetStatus():
    while True:
        try:
            global Btimer
            Btimer=browser.find_element('xpath','//clock')
            global Wily
            Wily=False
        except :
            pass
        else:
            break
    while True:
        try:
            Main_Ui.Vaziat_Bazaar.setText(Btimer.text)
        except Exception as e:
            try:
                WebDriverWait(browser, 5).until(EC.visibility_of_element_located(('xpath','//clock')))
                Btimer=browser.find_element('xpath','//clock')
                Main_Ui.Vaziat_Bazaar.setText(Btimer.text)
            except:
                pass

def NowTime(Element):
    Element.setTime(QTime.currentTime().addMSecs(CorrectionMSecs))

def PingServers():
    Main_Ui.Ping_btn.setEnabled(False)
    Server={0:'sptest2.asiatech.ir',1:'speedtest1.irancell.ir',2:'speedtest.pars.host',3:'speedtest1.shatel.ir'}
    os.system('start cmd /K ping '+Server[Main_Ui.PingCombo.currentIndex()])
    Main_Ui.Ping_btn.setEnabled(True)

def CorrectionMSecsAdder(msecs):
    global CorrectionMSecs
    CorrectionMSecs+=msecs
    Main_Ui.Correction_lbl.setText(' %s میلی ثانیه'%CorrectionMSecs)

def NTPCorrector():
    global CorrectionMSecs
    Main_Ui.NTP_btn.setEnabled(False)
    Main_Ui.AtomiClockCombo.setEnabled(False)
    palette = QtGui.QPalette()
    brush = QtGui.QBrush(QtGui.QColor(170, 116, 28))
    brush.setStyle(QtCore.Qt.SolidPattern)
    palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.WindowText, brush)
    brush = QtGui.QBrush(QtGui.QColor(170, 116, 28))
    brush.setStyle(QtCore.Qt.SolidPattern)
    palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.WindowText, brush)
    brush = QtGui.QBrush(QtGui.QColor(120, 120, 120))
    brush.setStyle(QtCore.Qt.SolidPattern)
    palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText, brush)
    Main_Ui.Correction_lbl_2.setPalette(palette)
    Main_Ui.Correction_lbl.setPalette(palette)
    Main_Ui.Correction_lbl.setText('در حال اتصال')
    Main_Ui.Correction_lbl_2.setText('در حال اتصال')
    try:
        time.sleep(0.5)
        call = ntplib.NTPClient()
        ntp_server=['ntp.day.ir','1.ir.pool.ntp.org','2.ir.pool.ntp.org','3.ir.pool.ntp.org','time.advtimesync.com','time.nist.gov']
        index=Main_Ui.AtomiClockCombo.currentIndex()
        response = call.request(ntp_server[index], version=3)
        servertime = datetime.fromtimestamp(response.orig_time)
        now=datetime.now()
        timeformat='%d/%m/%Y %H:%M:%S.%f'
        date1=now.strftime(timeformat)
        date2=servertime.strftime(timeformat)
        start = datetime.strptime(date1, timeformat)
        end =   datetime.strptime(date2, timeformat)
        diff = end - start
        CorrectionMSecs=math.ceil(diff.total_seconds() * 1000)
    except :

        Main_Ui.Correction_lbl_2.setText(' خطا')
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(170, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(170, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(120, 120, 120))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText, brush)
        Main_Ui.Correction_lbl_2.setPalette(palette)
    else:
        Main_Ui.Correction_lbl_2.setText(' دریافت شد')
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(2, 85, 2))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(2, 85, 2))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(120, 120, 120))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText, brush)
        Main_Ui.Correction_lbl_2.setPalette(palette)
    finally:
        Main_Ui.Correction_lbl.setText(' %s میلی ثانیه'%CorrectionMSecs)
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(2, 85, 2))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(2, 85, 2))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(120, 120, 120))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText, brush)
        Main_Ui.Correction_lbl.setPalette(palette)
        Main_Ui.NTP_btn.setEnabled(True)
        Main_Ui.AtomiClockCombo.setEnabled(True)

def OpenUs(type):
    match type:
        case 1:
            webbrowser.open('')
        case 2:
            webbrowser.open('')
        case 3:
            webbrowser.open('')

PageNameToHide = 'Sarkhati_09339496041'
Evalue=False
StopVal=False





if __name__ == "__main__"   :

    CorrectionMSecs=0
    global Thread4
    Thread4=threading.Thread(target=Hider)
    Thread4.start()
    global browser


    firefox_options = Options()
    firefox_options.add_argument("--start-maximized")
    firefox_options.set_preference("dom.webnotifications.enabled", False)
    firefox_options.set_preference("permissions.default.image", 2)
    firefox_options.set_preference("permissions.default.script", 2)
    firefox_options.set_preference("permissions.default.stylesheet", 2)
    firefox_options.set_preference("permissions.default.subdocument", 2)
    browser = webdriver.Chrome(options=firefox_options)
    #chrome_options.add_argument("--headless")
    browser.execute_script('document.title = "%s"' % PageNameToHide)
    actions = ActionChains(browser)
    TIMES = [1]
    MaxPrice = 0
    MinPrice = 0

    app = QtWidgets.QApplication(sys.argv)

    # Login_Form:
    # Login_Form = QtWidgets.QWidget()
    # Login_Ui = Login.Ui_Form()
    # Login_Ui.setupUi(Login_Form)
    # Login_Form.show()

    # Main_Form:
    Main_Ui = Main.Ui_Main_Form()
    Main_Form = QtWidgets.QWidget()
    Main_Ui.setupUi(Main_Form)
    Main_Form.show()
    time.sleep(1)
    Stop_hide()
    #Search Symbol Field Settings:
    Main_Ui.Search_Symbol.setFocusPolicy(QtCore.Qt.StrongFocus)
    Main_Ui.Search_Symbol.setEditable(True)
    Main_Ui.Search_Symbol.pFilterModel = QtCore.QSortFilterProxyModel(Main_Ui.Search_Symbol)
    Main_Ui.Search_Symbol.pFilterModel.setFilterCaseSensitivity(QtCore.Qt.CaseInsensitive)
    Main_Ui.Search_Symbol.pFilterModel.setSourceModel(Main_Ui.Search_Symbol.model())
    Main_Ui.Search_Symbol.completer = QtWidgets.QCompleter(Main_Ui.Search_Symbol.pFilterModel, Main_Ui.Search_Symbol)
    Main_Ui.Search_Symbol.completer.setCompletionMode(QtWidgets.QCompleter.UnfilteredPopupCompletion)

    # Time Settings:
    timer = QTimer()
    timer.timeout.connect(TimeView)
    timer.start()
    Main_Ui.TimeCorrector.clicked.connect( lambda: CorrectionMSecsAdder(Main_Ui.Correction.value()))
    Main_Ui.Start_Time.timeChanged.connect(lambda: Main_Ui.End_Time.setMinimumTime(Main_Ui.Start_Time.time()))
    Main_Ui.Now_btn.clicked.connect( lambda : NowTime(Main_Ui.Start_Time))
    Main_Ui.Now_btn_2.clicked.connect( lambda : NowTime(Main_Ui.End_Time))
    Main_Ui.Plus_btn.clicked.connect(lambda: Main_Ui.Start_Time.setTime(Main_Ui.Start_Time.time().addSecs(10)))
    Main_Ui.Plus_btn_2.clicked.connect(lambda: Main_Ui.End_Time.setTime(Main_Ui.End_Time.time().addSecs(10)))
    Main_Ui.Ping_btn.clicked.connect(PingServers)
    Main_Ui.NTP_btn.clicked.connect(lambda: threading.Thread(target=NTPCorrector).start())

    # Captcha_Form:
    Captcha_Ui = Captcha.Ui_Captcha_Form()
    Captcha_Form = QtWidgets.QWidget()
    Captcha_Ui.setupUi(Captcha_Form)
    Captcha_Form.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)
    # Login_Ui.Login.clicked.connect(lambda : [Main_Form.show(),Login_Form.close() , Stop_hide()])
    Main_Ui.Enter_btn.clicked.connect(Captcha_show)
    Captcha_Ui.Enter_btn.clicked.connect(Enter_site)
    Main_Ui.Search_Symbol.currentTextChanged.connect(lambda: [TIMES.append(time.time()),Disabler(1)])
    Thread1 = threading.Thread(target=Search_Symbol)
    Thread1.start()
    Main_Ui.Max_btn.clicked.connect(lambda: Main_Ui.Price_Field.setValue(MaxPrice))
    Main_Ui.Min_btn.clicked.connect(lambda: Main_Ui.Price_Field.setValue(MinPrice))
    Main_Ui.Quantity_Field.valueChanged.connect(lambda: Main_Ui.Total_Price.setText(str(Main_Ui.Quantity_Field.value() * Main_Ui.Price_Field.value())))
    Main_Ui.Price_Field.valueChanged.connect(lambda: Main_Ui.Total_Price.setText(str(Main_Ui.Quantity_Field.value() * Main_Ui.Price_Field.value())))    
    Main_Ui.SellOrBuy.currentIndexChanged.connect(lambda: Disabler(5))
    Main_Ui.Kargozar.currentIndexChanged.connect(lambda:[Main_Ui.User_Field.setEnabled(True),Main_Ui.Password_Field.setEnabled(True),Main_Ui.Enter_btn.setEnabled(True)])

    # Sarkhat Settings:
    Thread5 = threading.Thread(target=GetStatus)
    Thread5.start()
    Main_Ui.Sarkhat_btn.clicked.connect(lambda : [Starter(),Disabler(3)])
    Main_Ui.Stop_btn.clicked.connect(lambda: [Stopper(),Disabler(4)])

    #AboutUs
    # Main_Ui.AboutUs_btn.clicked.connect(lambda : OpenUs(1))
    # Main_Ui.Poshtibani_btn.clicked.connect(lambda : OpenUs(2))
    # Main_Ui.Pardazesh_btn.clicked.connect(lambda : OpenUs(3))



    os._exit(app.exec_())
else:
    os._exit(0)
    
