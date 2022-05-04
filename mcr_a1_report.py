import os
import pandas as pd
import xlwings as xw
from time import sleep
import pyautogui as gui
from selenium import webdriver
from selenium.webdriver import ActionChains



class DownloadMrcReport():

    def __init__(self):
        self.url = 'https://rb-wam-bi-p.wam-sso.bosch.com/BOE/BI/'
        self.source_path=r'\\bosch.com\dfsRB\dfsCN\LOC\Wx4\Dept\COR\03_CTG\80_IT tools\15_Python\Python MCR Index'
        self.save_path=r'\\bosch.com\dfsrb\DfsCN\loc\WX4\Dept\COR\03_CTG\80_IT tools\15_Python\R&D A1 Report'


    #进入到iframe里寻找元素，检查元素是否检查完毕
    def check_presence(self,browser,id,frame_list):
        try:
            for frame in frame_list:
                browser.switch_to.frame(frame)
            element=browser.find_element_by_id(id)
            return True
        except:
            sleep(2)
            browser.switch_to.default_content()
            self.check_presence(browser,id,frame_list)

    #检查出现在频幕上的特定图标是否已经加载完成,如果加载完成，则点击图标
    def check_gui_element_presence_click(self,el_path):
        points=gui.locateCenterOnScreen(el_path)
        if points:
            gui.click(points,button='left')
        else:
            sleep(2)
            self.check_gui_element_presence_click(el_path)

    def check_gui_element_presence(self,el_path):
        points=gui.locateCenterOnScreen(el_path)
        if points:
            return True
        else:
            sleep(180)
            self.check_gui_element_presence(el_path)

    def input_parameters(self,param):
        gui.typewrite(param)
        sleep(1)
        gui.press('tab')
        sleep(1)
        gui.press('tab')
        sleep(1)

    def tab_down(self,num):
        for i in range(num):
            gui.press('tab')
            sleep(0.3)

    def download_report(self):
        name_time=str(pd.datetime.now()).split(' ')[0].replace('-','_')
        doc_params=pd.read_excel(self.source_path+os.sep+'MCR_A1_version_RPA.xlsx')
        doc_params['LO version']=doc_params['LO version'].astype('str')
        doc_params['SP version']=doc_params['SP version'].astype('str')
        lo_data=doc_params['LO version'].values.tolist()
        sp_data=doc_params['SP version'].values.tolist()
        parms_data= {'lo':[lo_data,'R11_A1_LO-View_Legal_Share_Consolidated_RBCW_{time}_EUR.xlsx'.format(time=name_time),'ListingURE_detailView_listNode0_0'],'sp':[sp_data,'R11_A1_SP-View_Legal_Share_Consolidated_RBCW_{time}_EUR.xlsx'.format(time=name_time),'ListingURE_detailView_listNode1_0']}

        for check_point in ['lo','sp']:
            browser = webdriver.Firefox()
            browser.get(self.url)
            # 等待首页加载完毕，点击documents标签
            self.check_presence(browser, 'id_56', ['servletBridgeIframe', 'iframeHome-23789649'])
            browser.switch_to.default_content()
            browser.switch_to.frame('servletBridgeIframe')
            document_tab = browser.find_element_by_css_selector('#id_8 > div:nth-child(2)')
            document_tab.click()
            print('Click document tab')
            # 等待Documents加载完毕，点击第一个文件链接
            self.check_presence(browser, 'ListingURE_detailView_listNode0_0',['servletBridgeIframe', 'iframe4424-23789649'])
            first_file=browser.find_element_by_id(parms_data[check_point][2])
            actions=ActionChains(browser)
            #点击第一个链接
            actions.double_click(first_file).perform()
            print('CLick file link')
            #点击ok按钮
            self.check_gui_element_presence_click(self.source_path+os.sep+'ok_button.png')
            print('Click OK button')
            #等待prj标签并点击
            self.check_gui_element_presence_click(self.source_path+os.sep+'prj_division_button.png')
            print('Wait pjr division label and start to input params')
            self.input_parameters(parms_data[check_point][0][0])
            gui.typewrite(parms_data[check_point][0][1])
            sleep(2)
            self.tab_down(28)
            sleep(1)
            #等待currency标签并点击
            self.check_gui_element_presence_click(self.source_path+os.sep+'currency_button.png')
            sleep(1)
            gui.press(['delete']*5)
            sleep(3)
            for count in range(13):
                self.input_parameters(parms_data[check_point][0][count+2])
            gui.press(['tab']*2)
            sleep(0.5)
            gui.press('space')
            print('Start to download data from server and wait')
            sleep(15)
            self.check_gui_element_presence(self.source_path+os.sep+'search_button.png')
            print('start to copy data')
            #book是源文件，book1是新建的文件，用来存储复制的数据
            book=xw.books.active
            sht=book.sheets('Tabelle1')
            #复制数据
            copy_values=sht.used_range.value
            #新建工作表
            app1=xw.App(visible=False,add_book=False)
            book1=app1.books.add()
            sht1=book1.sheets('Sheet1')
            print('Paste data')
            sht1.range('A1').value=copy_values
            sht1.autofit('c')
            if os.path.exists(self.save_path+os.sep+parms_data[check_point][1]):
                #如果文件存在，则删除文件
                os.remove(self.save_path+os.sep+parms_data[check_point][1])
            book1.save(self.save_path+os.sep+parms_data[check_point][1])
            book1.close()
            app1.quit()
            app = xw.apps.active
            app.quit()
            print('Close Excel Application')
            print('------------------------------------')
            browser.quit()
d=DownloadMrcReport()
d.download_report()
