import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import UnexpectedAlertPresentException
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import xlsxwriter
import xlrd

def excel_column_name(n):
    """Number to Excel-style column name, e.g., 1 = A, 26 = Z, 27 = AA, 703 = AAA."""
    name = ''
    while n > 0:
        n, r = divmod (n - 1, 26)
        name = chr(r + ord('A')) + name
    return name

def crawling_recent_trade(read_data, month_select_op):
        
    options = Options()
    options.headless = True
    #options.headless = False
    browser = webdriver.Chrome(executable_path="./chromedriver.exe", options=options)

    #print(read_data)

    browser.get("http://stat.kita.net/stat/kts/prod/ProdWholeDetailPopup.screen#none")
    print("Get URL")

    try:
        result = browser.switch_to_alert()
        print(result.text)
        result.accept()
        result.dismiss()
    except:
        print("There is no alert")

    crawling_list = []
    wait = WebDriverWait(browser, 40)
    error_list = []

    for n in range(len(read_data)):


        prod_code = read_data[n][0]
        item_value = read_data[n][2]
        
        print("data index : ", n, prod_code, item_value)

        time.sleep(0.5)

        s_prod_code = browser.find_element_by_css_selector("#s_prod_code")
        s_prod_code.clear()
        s_prod_code.send_keys(prod_code)
       
        time.sleep(0.5)

        s_item_value = browser.find_element_by_css_selector("#s_item_value")
        s_item_value.clear()
        s_item_value.send_keys(item_value)

        time.sleep(0.5)

        monthsum_select = Select(browser.find_element_by_css_selector("body > form > div > div > fieldset > div:nth-child(2) > div:nth-child(2) > select"))
        monthsum_select.select_by_value('1')

        time.sleep(0.5)

        scale_select = Select(browser.find_element_by_css_selector("body > form > div > div > fieldset > div:nth-child(2) > div:nth-child(3) > select"))
        scale_select.select_by_value('1000')

        browser.find_element_by_xpath("/html/body/form/div/div/fieldset/div[3]/a/img").click()

        #time.sleep(10)

        print("Expand Recent")

        try:

            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#mySheet1 > tbody > tr:nth-child(3) > td > div > div.GMPageOne > table > tbody > tr:nth-child(2) > td.GMClassReadOnly.GMCell.GMNoRight.GMEmpty.GME.IBSheetFont0')))
            time.sleep(0.5)

            #try:
            browser.find_element_by_css_selector("#mySheet1 > tbody > tr:nth-child(3) > td > div > div.GMPageOne > table > tbody > tr:nth-child(2) > td.GMClassReadOnly.GMCell.GMNoRight.GMEmpty.GME.IBSheetFont0").click()
            #except:
            
            time.sleep(0.5)
            #browser.implicitly_wait(20)
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#mySheet1 > tbody > tr:nth-child(3) > td > div > div.GMPageOne > table > tbody > tr:nth-child(3) > td > table > tbody > tr:nth-child(2) > td.GMClassReadOnly.GMWrap0.GMAlignLeft.GMText.GMCell.IBSheetFont0.GMNoLeft.HideCol0C1')))                                                               
                        
            html = browser.page_source
            soup = BeautifulSoup(html, 'html.parser', from_encoding='utf-8')
            #print(soup)
            trs = soup.find_all('tr', {"class":"GMDataRow"})
            #for tr in trs:
            #    print(tr.text)

            # len 17
            # ['', '\xa0', '2020년', '36,631,568', '1,934.1', '21,961,664', '169.9', '0', '0.0', '32,455,714', '-11.0', '239,694,362', '-4.8', '0', '0.0', '4,175,854', '2020']
            recent_tr_list = []

            idx_recent_tr = 0
            for k in range(1, len(trs)):
                tds = trs[k].find_all('td')
                if tds[2].text.find("년") == -1:
                    idx_recent_tr = k
                    #print("idx : ", idx_recent_tr, tds[2].text)

            tds = trs[idx_recent_tr].find_all('td')
            tds_next = trs[idx_recent_tr-1].find_all('td')
            td_list = []

            #년월           td[2]
            #수출 금액      td[3]
            #수출 증감률    td[4]
            #수출 중량      td[5]
            #수출 증감률    td[6]
            #수출 금액 MoM  td[7]
            #수출 중량 MoM  td[8]
            #수입 금액      td[9]
            #수입 증감률    td[10]
            #수입 중량      td[11]
            #수입 증감률    td[12]
            #수입 금액 MoM  td[13]
            #수입 중량 MoM  td[14]
            #수지           td[15]

            str_year = "2020년"

            for k in range(len(tds)):
                if k==2:
                    if tds[k].text.find("년") == -1:
                        #print(tds[k].text)
                        td_list.append(str_year + tds[k].text)            
                    else:
                        str_year = tds[k].text
                        td_list.append(tds[k].text)
                elif k==4 or k==6 or k==10 or k==12:
                    td_list.append(float(tds[k].text.replace(',','')))
                elif k==3 or k==5 or k==9 or k==11 or k==15:
                    td_list.append(int(tds[k].text.replace(',','')))
                elif k==7:
                    #print(int(tds[3].text.replace(',','')), int(tds_next[3].text.replace(',','')))
                    if int(tds_next[3].text.replace(',','')) == 0:
                        td_list.append(0.0)
                    else:
                        mom = (int(tds[3].text.replace(',',''))-int(tds_next[3].text.replace(',','')))/float(tds_next[3].text.replace(',',''))*100
                        td_list.append(float(mom))
                elif k==8:
                    #print(int(tds[5].text.replace(',','')), int(tds_next[5].text.replace(',','')))
                    if int(tds_next[5].text.replace(',','')) == 0:
                        td_list.append(0.0)
                    else:
                        mom = (int(tds[5].text.replace(',',''))-int(tds_next[5].text.replace(',','')))/float(tds_next[5].text.replace(',',''))*100
                        td_list.append(float(mom))
                elif k==13:
                    if int(tds_next[9].text.replace(',','')) == 0:
                        td_list.append(0.0)
                    else:
                        mom = (int(tds[9].text.replace(',',''))-int(tds_next[9].text.replace(',','')))/float(tds_next[9].text.replace(',',''))*100
                        td_list.append(float(mom))
                elif k==14:
                    if int(tds_next[11].text.replace(',','')) == 0:
                        td_list.append(0.0)
                    else:
                        mom = (int(tds[11].text.replace(',',''))-int(tds_next[11].text.replace(',','')))/float(tds_next[11].text.replace(',',''))*100
                        td_list.append(float(mom))
                else:
                    td_list.append(tds[k].text)

            recent_tr_list.append(td_list)

            crawling_list.append(recent_tr_list)   

        except:
            recent_tr_list = []
            recent_tr_list.append([0,0,'0',0,0,0,0,0,0,0,0,0,0,0,0,0])
            crawling_list.append(recent_tr_list)
            error_list.append([prod_code, item_value])
        
    browser.close()
    print(error_list)

    return crawling_list


def crawling_all_trade(read_data, month_select_op):

    options = Options()
    #options.headless = True
    options.headless = False
    browser = webdriver.Chrome(executable_path="./chromedriver.exe", options=options)

    print(read_data)

    #url = "http://stat.kita.net/stat/kts/prod/ProdWholeDetailPopup.screen#none"
    #browser.get("http://stat.kita.net/stat/kts/prod/ProdWholeList.screen")
    browser.get("http://stat.kita.net/stat/kts/prod/ProdWholeDetailPopup.screen#none")
    print("Get URL")
    #browser.wait_for_and_accept_alert()

    #browser.get(url)
    try:
    #    browser.get(url)
        #browser.get("http://stat.kita.net/stat/kts/prod/ProdWholeDetailPopup.screen#none")
        result = browser.switch_to_alert()
        print(result.text)
        result.accept()
        result.dismiss()
    except:
    #except UnexpectedAlertPresentException:
        #alert = browser.switch_to_alert()
        #alert.dismiss()
        print("There is no alert")


    # browser.get("", data)

    # s_term_gb
    # s_monthsum_gb
    # s_measure
    # s_prod_name
    # s_year

    crawling_list = []

    #prod_code = "460"
    #item_value = "382200"

    wait = WebDriverWait(browser, 40)

    for n in range(len(read_data)):

        prod_code = read_data[n][0]
        item_value = read_data[n][2]
        
        print("data index : ", n, prod_code, item_value)

        time.sleep(0.5)

        s_prod_code = browser.find_element_by_css_selector("#s_prod_code")
        s_prod_code.clear()
        s_prod_code.send_keys(prod_code)
       

        time.sleep(0.5)

        s_item_value = browser.find_element_by_css_selector("#s_item_value")
        #for i in range(5):
        #    s_item_value.send_keys(Keys.BACKSPACE)
        #    time.sleep(0.5)
        s_item_value.clear()
        s_item_value.send_keys(item_value)

        time.sleep(0.5)

        monthsum_select = Select(browser.find_element_by_css_selector("body > form > div > div > fieldset > div:nth-child(2) > div:nth-child(2) > select"))
        monthsum_select.select_by_value('1')

        time.sleep(0.5)

        scale_select = Select(browser.find_element_by_css_selector("body > form > div > div > fieldset > div:nth-child(2) > div:nth-child(3) > select"))
        scale_select.select_by_value('1000')

        browser.find_element_by_xpath("/html/body/form/div/div/fieldset/div[3]/a/img").click()

        time.sleep(1)
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#mySheet1 > tbody > tr:nth-child(3) > td > div > div.GMPageOne > table > tbody > tr:nth-child(2) > td.GMClassReadOnly.GMCell.GMNoRight.GMEmpty.GME.IBSheetFont0')))

        if month_select_op == 1:

            ## Scroll DOWN
            try:
                browser.find_element_by_css_selector("#mySheet1 > tbody > tr:nth-child(2) > td:nth-child(2) > div > div").click()
                actions = ActionChains(browser)
                for i in range(3):
                    actions.send_keys(Keys.PAGE_DOWN).perform()
                    time.sleep(0.5)
            except:
                print("Scroll Bar error")

            #mySheet1 > tbody > tr:nth-child(2) > td:nth-child(2) > div > div

            click_interval = 3
            

            tag_names = browser.find_element_by_css_selector("#mySheet1 > tbody > tr:nth-child(3) > td > div > div.GMPageOne > table").find_elements_by_tag_name("tr")
            #for tag in tag_names:
            #    print (tag.text.split("\n"))
            
            tr_cnt = len(tag_names)
            print(tr_cnt)

            print("Expand")

            browser.find_element_by_css_selector("#mySheet1 > tbody > tr:nth-child(3) > td > div > div.GMPageOne > table > tbody > tr:nth-child(" + str(tr_cnt)+") > td.GMClassReadOnly.GMCell.GMNoRight.GMEmpty.GMEL.IBSheetFont0").click()
            #try:
            #    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#mySheet1 > tbody > tr:nth-child(3) > td > div > div.GMPageOne > table > tbody > tr:nth-child(" + str(tr_cnt+1) + ") > td > table > tbody > tr:nth-child(2) > td.GMClassReadOnly.GMWrap0.GMAlignLeft.GMText.GMCell.IBSheetFont0.GMNoLeft.HideCol0C1")))         
            #except:
            #    print("First Expand Error")
            time.sleep(3.0)    
            for idx in range(tr_cnt-2):
                print("Expand Index : ", idx, "str(tr_cnt-2-idx) ", tr_cnt-2-idx)
                browser.find_element_by_css_selector("#mySheet1 > tbody > tr:nth-child(3) > td > div > div.GMPageOne > table > tbody > tr:nth-child(" + str(tr_cnt-1-idx) + ") > td.GMClassReadOnly.GMCell.GMNoRight.GMEmpty.GME.IBSheetFont0").click()
                wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#mySheet1 > tbody > tr:nth-child(3) > td > div > div.GMPageOne > table > tbody > tr:nth-child(" + str(tr_cnt-idx) + ") > td > table > tbody > tr:nth-child(2) > td.GMClassReadOnly.GMWrap0.GMAlignLeft.GMText.GMCell.IBSheetFont0.GMNoLeft.HideCol0C1")))         
                #time.sleep(click_interval)

                if idx == 12:            
                    #Scroll UP
                    actions = ActionChains(browser)
                    for i in range(3):
                        actions.send_keys(Keys.PAGE_UP).perform()
                        time.sleep(0.5)

        html = browser.page_source
        soup = BeautifulSoup(html, 'html.parser', from_encoding='utf-8')
        #print(soup)
        trs = soup.find_all('tr', {"class":"GMDataRow"})

        # Write BS to TXT file
        #f = open("./test.txt", "w", encoding='utf-8')
        #f.write(str(soup))
        #f.close()


        # len 17
        # ['', '\xa0', '2020년', '36,631,568', '1,934.1', '21,961,664', '169.9', '0', '0.0', '32,455,714', '-11.0', '239,694,362', '-4.8', '0', '0.0', '4,175,854', '2020']
        tr_list = []

        str_year = "2030년"
        
        for tr in trs:
            
            pass_tr = 0
            
            #print(tr.text.split("\n"))
            #print(tr.text)
            tds = tr.find_all('td')
            td_list = []

            #년월           td[2]
            #수출 금액      td[3]
            #수출 증감률    td[4]
            #수출 중량      td[5]
            #수출 증감률    td[6]
            #수입 금액      td[9]
            #수입 증감률    td[10]
            #수입 중량      td[11]
            #수입 증감률    td[12]
            #수지           td[15]

            for k in range(len(tds)):
                if k==2:
                    if tds[k].text.find("년") == -1:
                        #print(tds[k].text)
                        td_list.append(str_year + tds[k].text)            
                    else:
                        str_year = tds[k].text
                        td_list.append(tds[k].text)
                        if month_select_op == 1:
                            pass_tr = 1;
                elif k==4 or k==6 or k==10 or k==12:
                    td_list.append(float(tds[k].text.replace(',','')))
                elif k==3 or k==5 or k==9 or k==11 or k==15:
                    td_list.append(int(tds[k].text.replace(',','')))
                else:
                    td_list.append(tds[k].text)

            #for td in tds:
            #    td_list.append(td.text)
            #    #print(td.text)

            #rint("len", len(td_list))
            if pass_tr == 0:
                tr_list.append(td_list)

     
        #print(tr_list)

        #print("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$")
        #browser.find_element_by_css_selector("#mySheet1 > tbody > tr:nth-child(3) > td > div > div.GMPageOne > table > tbody > tr:nth-child(5) > td > table > tbody > tr:nth-child(10) > td.GMClassReadOnly.GMWrap0.GMAlignLeft.GMText.GMCell.IBSheetFont0.GMNoLeft.HideCol0C1").click()

        '''
        for k in range(3):
            print("=================================================")
            for j in range(15):
                actions.send_keys(Keys.ARROW_DOWN).perform()
                
            print(browser.find_element_by_css_selector("#mySheet1").text)
            time.sleep(1)


        time.sleep(3)
        tag_names = browser.find_element_by_css_selector("#mySheet1").find_elements_by_tag_name("tr")
        for tag in tag_names:
            print (tag.text.split("\n"))

        '''

        #print(browser.find_element_by_css_selector("#mySheet1 > tbody > tr:nth-child(3)").text)
        
        if month_select_op == 1:
            tr_list.reverse()
            new_tr_list = []
            idx = 0
            
            while idx < len(tr_list):
                if idx + 12 > len(tr_list):
                    temp_list = tr_list[idx:]
                    temp_list.reverse()
                    new_tr_list = new_tr_list + temp_list
                    #print(new_tr_list)
                else:
                    #print(type(new_tr_list), type(tr_list[idx:idx+12]), type(tr_list[idx:idx+12].reverse()))
                    temp_list = tr_list[idx:idx+12]
                    temp_list.reverse()
                    new_tr_list = new_tr_list + temp_list
                    #print(new_tr_list)
                idx = idx + 12

            crawling_list.append(new_tr_list)      
        else:
            tr_list.reverse()
            crawling_list.append(tr_list)

    browser.close()

    return crawling_list

def write_excel_file(read_data, result_list, recent_op, view_graph_op):

    workbook_name = "trade_test.xlsx"
    #workbook_name = "trade_test_example.xlsx"
    workbook = xlsxwriter.Workbook(workbook_name)

    filter_format = workbook.add_format({'bold':True, 'fg_color': '#D7E4BC'	})
    filter_format.set_border()
    filter_format2 = workbook.add_format({'bold':True })
    filter_format2.set_border()
    filter_format3 = workbook.add_format({})
    filter_format3.set_border()

    percent_format = workbook.add_format({'num_format': '0.00%'})
    num_format = workbook.add_format({'num_format':'0.00'})
    num_format.set_border()
    num2_format = workbook.add_format({'num_format':'#,##0'})
    num2_format.set_border()
    #num3_format = workbook.add_format({'num_format':'#,##0.00', 'fg_color':'#FCE4D6'})

    graph_inc_list = []
    graph_dec_list = []

    if recent_op == 1:
        
        worksheet_name ='최근월 수출입통계'
        worksheet0 = workbook.add_worksheet(worksheet_name)
        worksheet0.set_column(0,0,20)

        for n in range(len(result_list)):

            result_sheet = result_list[n]

            prod_code = read_data[n][0]
            prod_name = read_data[n][1]
            item_value = read_data[n][2]
            item_name = read_data[n][3]
            corp_name = read_data[n][4]

            worksheet0.write(1, 0, "품목", filter_format3)
            worksheet0.write(1, 1, "년월", filter_format3)
            worksheet0.write(0, 2, "수출", filter_format3)
            worksheet0.write(1, 2, "금액", filter_format3)
            worksheet0.write(1, 3, "증감률(YoY)", filter_format3)
            worksheet0.write(1, 4, "증감률(MoM)", filter_format3)
            worksheet0.write(1, 5, "중량", filter_format3)
            worksheet0.write(1, 6, "증감률(YoY)", filter_format3)
            worksheet0.write(1, 7, "증감률(MoM)", filter_format3)
            worksheet0.write(0, 8, "수입", filter_format3)
            worksheet0.write(1, 8, "금액", filter_format3)
            worksheet0.write(1, 9, "증감률(YoY)", filter_format3)
            worksheet0.write(1, 10, "증감률(YoY)", filter_format3)
            worksheet0.write(1, 11, "중량", filter_format3)
            worksheet0.write(1, 12, "증감률(YoY)", filter_format3)
            worksheet0.write(1, 13, "증감률(YoY)", filter_format3)
            worksheet0.write(1, 14, "수지", filter_format3)
            worksheet0.write(1, 15, "관련회사", filter_format3)
            worksheet0.write(1, 16, "품목", filter_format3)

            offset = 2
            
            for i in range(len(result_sheet)):
                #print(i)
                #print("result_sheet", result_sheet)
                #print(result_sheet[i])
                # 분류
                worksheet0.write(n+offset, 0, str(prod_code) + " " + str(prod_name) + " " + str(item_value), filter_format)
                # 년월
                #print(result_sheet[i][2])
                worksheet0.write(n+offset,1, result_sheet[i][2], filter_format3)
                # 수출
                worksheet0.write(n+offset,2, result_sheet[i][3], num2_format)
                worksheet0.write(n+offset,3, result_sheet[i][4], num_format)
                worksheet0.write(n+offset,4, result_sheet[i][7], num_format)
                worksheet0.write(n+offset,5, result_sheet[i][5], num2_format)
                worksheet0.write(n+offset,6, result_sheet[i][6], num_format)
                worksheet0.write(n+offset,7, result_sheet[i][8], num_format)
                # 수입
                worksheet0.write(n+offset,8, result_sheet[i][9], num2_format)
                worksheet0.write(n+offset,9, result_sheet[i][10], num_format)
                worksheet0.write(n+offset,10, result_sheet[i][13], num_format)
                worksheet0.write(n+offset,11, result_sheet[i][11], num2_format)
                worksheet0.write(n+offset,12, result_sheet[i][12], num_format)
                worksheet0.write(n+offset,13, result_sheet[i][14], num_format)
                # 수지
                worksheet0.write(n+offset,14, result_sheet[i][15], num2_format)
                #관련회사
                worksheet0.write(n+offset,15, str(corp_name), filter_format3)
                #품목
                worksheet0.write(n+offset,16, str(item_name), filter_format3)

                if result_sheet[i][4] > 20.0 and result_sheet[i][7] > 20.0 and result_sheet[i][3] > 10000:
                    graph_inc_list.append([str(item_value), result_sheet[i][3], result_sheet[i][4], result_sheet[i][7]])
                elif result_sheet[i][4] < -20.0 and result_sheet[i][7] < -20.0 and result_sheet[i][3] > 10000:
                    graph_dec_list.append([str(item_value), result_sheet[i][3], result_sheet[i][4], result_sheet[i][7]])

        if view_graph_op == 1:
            
            worksheet1 = workbook.add_worksheet("Graph")
            offset = 3

            dec_list_len = len(graph_dec_list)
            #char_dec = chr(65+dec_list_len)
            inc_list_len = len(graph_inc_list)
            #char_inc = chr(65+inc_list_len)

            print(len(graph_dec_list))
            print(len(graph_inc_list))

            worksheet1.write(offset, 0, "품목", filter_format3)
            worksheet1.write(offset+1, 0, "금액", filter_format3)
            worksheet1.write(offset+2, 0, "증감률(YoY)", filter_format3)
            worksheet1.write(offset+3, 0, "증감률(MoM)", filter_format3)

            worksheet1.write(offset+5, 0, "품목", filter_format3)
            worksheet1.write(offset+6, 0, "금액", filter_format3)
            worksheet1.write(offset+7, 0, "증감률(YoY)", filter_format3)
            worksheet1.write(offset+8, 0, "증감률(MoM)", filter_format3)

            for k in range(len(graph_inc_list)):
                worksheet1.write(offset,k+1, graph_inc_list[k][0])
                worksheet1.write(offset+1,k+1, graph_inc_list[k][1])
                worksheet1.write(offset+2,k+1, graph_inc_list[k][2])
                worksheet1.write(offset+3,k+1, graph_inc_list[k][3])

            for k in range(len(graph_dec_list)):
                worksheet1.write(offset+5,k+1, graph_dec_list[k][0])
                worksheet1.write(offset+6,k+1, graph_dec_list[k][1])
                worksheet1.write(offset+7,k+1, graph_dec_list[k][2])
                worksheet1.write(offset+8,k+1, graph_dec_list[k][3])

            #bar_chart1 = workbook.add_chart({'type':'bar', 'subtype':'stacked'})
            bar_chart1 = workbook.add_chart({'type':'column'})

            bar_chart1.add_series({
                'name': '증감률(YoY)',
                'categories': 'Graph!$B$4:$' + excel_column_name(inc_list_len+1) + '$4',
                'values': 'Graph!$B$6:$' + excel_column_name(inc_list_len+1) + '$6',
                #'y2_axis': True,
            })
            bar_chart1.add_series({
                'name': '증감률(MoM)',
                'categories': 'Graph!$B$4:$' + excel_column_name(inc_list_len+1) + '$4',
                'values': 'Graph!$B$7:$' + excel_column_name(inc_list_len+1) + '$7',
                #'y2_axis': True,
            })

            bar_chart1.set_size({'width':600, 'height':400})
            bar_chart1.set_title({'name':'수출 증가'})
            bar_chart1.set_x_axis({'name': '품목'})
            bar_chart1.set_y_axis({'name':'%'})
            bar_chart1.set_legend({'position':'bottom'})

            worksheet1.insert_chart('B16', bar_chart1)

            #bar_chart2 = workbook.add_chart({'type':'bar', 'subtype':'stacked'})
            bar_chart2 = workbook.add_chart({'type':'column'})
            bar_chart2.add_series({
                'name': '증감률(YoY)',
                'categories': 'Graph!$B$9:$' + excel_column_name(dec_list_len) + '$9',
                'values': 'Graph!$B$11:$' + excel_column_name(dec_list_len) + '$11',
                #'y2_axis': True,
            })
            bar_chart2.add_series({
                'name': '증감률(MoM)',
                'categories': 'Graph!$B$9:$' + excel_column_name(dec_list_len) + '$9',
                'values': 'Graph!$B$12:$' + excel_column_name(dec_list_len) + '$12',
                #'y2_axis': True,
            })

            bar_chart2.set_size({'width':600, 'height':400})
            bar_chart2.set_title({'name':'수출 감소'})
            bar_chart2.set_x_axis({'name': '품목'})
            bar_chart2.set_y_axis({'name':'%'})
            bar_chart2.set_legend({'position':'bottom'})

            worksheet1.insert_chart('L16', bar_chart2)
            

    else:
        for n in range(len(result_list)):
            
            prod_code = read_data[n][0]
            prod_name = read_data[n][1]
            item_value = read_data[n][2]
            item_name = read_data[n][3]
            corp_name = read_data[n][4]

            worksheet_name = str(prod_name).replace(' ','') + "_" + str(item_value).replace(' ','')
            worksheet0 = workbook.add_worksheet(worksheet_name)

            result_sheet = result_list[n]

            #년월           td[2]
            #수출 금액      td[3]
            #수출 증감률    td[4]
            #수출 중량      td[5]
            #수출 증감률    td[6]
            #수입 금액      td[9]
            #수입 증감률    td[10]
            #수입 중량      td[11]
            #수입 증감률    td[12]
            #수지           td[15]

            worksheet0.write(0, 0, str(prod_code) + " " + str(prod_name) + " " + str(item_value) + " " + item_name + " " + corp_name, filter_format)
            worksheet0.set_column(0,0,20)

            worksheet0.write(2, 0, "년월", filter_format3)
            worksheet0.write(1, 1, "수출", filter_format3)
            worksheet0.write(2, 1, "금액", filter_format3)
            worksheet0.write(2, 2, "증감률", filter_format3)
            worksheet0.write(2, 3, "중량", filter_format3)
            worksheet0.write(2, 4, "증감률", filter_format3)
            worksheet0.write(1, 5, "수입", filter_format3)
            worksheet0.write(2, 5, "금액", filter_format3)
            worksheet0.write(2, 6, "증감률", filter_format3)
            worksheet0.write(2, 7, "중량", filter_format3)
            worksheet0.write(2, 8, "증감률", filter_format3)
            worksheet0.write(2, 9, "수지", filter_format3)

            offset = 3
            row_idx = len(result_sheet) + 3

            for i in range(len(result_sheet)):
                # 년월
                worksheet0.write(i+offset,0, result_sheet[i][2], filter_format3)
                # 수출
                worksheet0.write(i+offset,1, result_sheet[i][3], num2_format)
                worksheet0.write(i+offset,2, result_sheet[i][4], num_format)
                worksheet0.write(i+offset,3, result_sheet[i][5], num2_format)
                worksheet0.write(i+offset,4, result_sheet[i][6], num_format)
                # 수입
                worksheet0.write(i+offset,5, result_sheet[i][9], num2_format)
                worksheet0.write(i+offset,6, result_sheet[i][10], num_format)
                worksheet0.write(i+offset,7, result_sheet[i][11], num2_format)
                worksheet0.write(i+offset,8, result_sheet[i][12], num_format)
                # 수지
                worksheet0.write(i+offset,9, result_sheet[i][15], num2_format)

                #for j in range(len(result_sheet[i])):
                #    worksheet0.write(i+1,j, result_sheet[i][j])

            # Chart
            column_chart = workbook.add_chart({'type':'column'})
            column_chart.add_series({
                'name': '수출금액',
                'categories': worksheet_name + '!$A$4:$A$' + str(row_idx),
                'values': worksheet_name + '!$B$4:$B$' + str(row_idx),
                #'y2_axis': True,
            })

            #column_chart.combine(line_chart)

            column_chart.set_size({'width':600, 'height':400})
            column_chart.set_title({'name':'수출 금액'})
            column_chart.set_x_axis({'name': '년월'})
            column_chart.set_y_axis({'name':'천불'})
            column_chart.set_legend({'position':'bottom'})

            worksheet0.insert_chart('M6', column_chart)

            line_chart = workbook.add_chart({'type':'line'})
            line_chart.add_series({
                'name': '수출금액 YoY',
                'categories': worksheet_name + '!$A$4:$A$' + str(row_idx),
                'values': worksheet_name + '!$C$4:$C$' + str(row_idx)
            })

            line_chart.set_size({'width':600, 'height':400})
            line_chart.set_title({'name':'수출 금액 YoY'})
            line_chart.set_x_axis({'name': '년월'})
            line_chart.set_y_axis({'name':'%'})
            line_chart.set_legend({'position':'bottom'})
            
            worksheet0.insert_chart('M28', line_chart)

            column_chart = workbook.add_chart({'type':'column'})
            column_chart.add_series({
                'name': '수출금액',
                'categories': worksheet_name + '!$A$4:$A$' + str(row_idx),
                'values': worksheet_name + '!$D$4:$D$' + str(row_idx),
                #'y2_axis': True,
            })

            #column_chart.combine(line_chart)

            column_chart.set_size({'width':600, 'height':400})
            column_chart.set_title({'name':'수출 중량'})
            column_chart.set_x_axis({'name': '년월'})
            column_chart.set_y_axis({'name':'천불'})
            column_chart.set_legend({'position':'bottom'})

            worksheet0.insert_chart('W6', column_chart)

            line_chart = workbook.add_chart({'type':'line'})
            line_chart.add_series({
                'name': '수출금액 YoY',
                'categories': worksheet_name + '!$A$4:$A$' + str(row_idx),
                'values': worksheet_name + '!$E$4:$E$' + str(row_idx)
            })

            line_chart.set_size({'width':600, 'height':400})
            line_chart.set_title({'name':'수출 중량 YoY'})
            line_chart.set_x_axis({'name': '년월'})
            line_chart.set_y_axis({'name':'%'})
            line_chart.set_legend({'position':'bottom'})
            
            worksheet0.insert_chart('W28', line_chart)

            # Chart
            column_chart = workbook.add_chart({'type':'column'})
            column_chart.add_series({
                'name': '수입금액',
                'categories': worksheet_name + '!$A$4:$A$' + str(row_idx),
                'values': worksheet_name + '!$F$4:$F$' + str(row_idx),
                #'y2_axis': True,
            })

            #column_chart.combine(line_chart)

            column_chart.set_size({'width':600, 'height':400})
            column_chart.set_title({'name':'수입 금액'})
            column_chart.set_x_axis({'name': '년월'})
            column_chart.set_y_axis({'name':'천불'})
            column_chart.set_legend({'position':'bottom'})

            worksheet0.insert_chart('M50', column_chart)

            line_chart = workbook.add_chart({'type':'line'})
            line_chart.add_series({
                'name': '수입금액 YoY',
                'categories': worksheet_name + '!$A$4:$A$' + str(row_idx),
                'values': worksheet_name + '!$G$4:$G$' + str(row_idx)
            })

            line_chart.set_size({'width':600, 'height':400})
            line_chart.set_title({'name':'수입 금액 YoY'})
            line_chart.set_x_axis({'name': '년월'})
            line_chart.set_y_axis({'name':'%'})
            line_chart.set_legend({'position':'bottom'})
            
            worksheet0.insert_chart('M72', line_chart)

            column_chart = workbook.add_chart({'type':'column'})
            column_chart.add_series({
                'name': '수입금액',
                'categories': worksheet_name + '!$A$4:$A$' + str(row_idx),
                'values': worksheet_name + '!$H$4:$H$' + str(row_idx),
                #'y2_axis': True,
            })

            #column_chart.combine(line_chart)

            column_chart.set_size({'width':600, 'height':400})
            column_chart.set_title({'name':'수입 중량'})
            column_chart.set_x_axis({'name': '년월'})
            column_chart.set_y_axis({'name':'천불'})
            column_chart.set_legend({'position':'bottom'})

            worksheet0.insert_chart('W50', column_chart)

            line_chart = workbook.add_chart({'type':'line'})
            line_chart.add_series({
                'name': '수입금액 YoY',
                'categories': worksheet_name + '!$A$4:$A$' + str(row_idx),
                'values': worksheet_name + '!$I$4:$I$' + str(row_idx)
            })

            line_chart.set_size({'width':600, 'height':400})
            line_chart.set_title({'name':'수입 중량 YoY'})
            line_chart.set_x_axis({'name': '년월'})
            line_chart.set_y_axis({'name':'%'})
            line_chart.set_legend({'position':'bottom'})
            
            worksheet0.insert_chart('W72', line_chart)

    workbook.close()

def read_req_excel_file(input_file):

    read_data = []

    workbook_name = input_file
    workbook = xlrd.open_workbook(workbook_name)
    sheet_list = workbook.sheets()
    sheet1 = sheet_list[0]

    num_req = int(sheet1.cell(0,0).value)

    for i in range(num_req):
        region_code = int(sheet1.cell(i+2,0).value)
        region_name = sheet1.cell(i+2,1).value
        #item_code = int(sheet1.cell(i+2,2).value)
        item_code = sheet1.cell(i+2,2).value
        item_name = sheet1.cell(i+2,3).value
        corp_name = sheet1.cell(i+2,4).value

        print(region_code, region_name, item_code)
        read_data.append([region_code, region_name, item_code, item_name, corp_name])
       
    return read_data

def main():

    #[TODO] 
    # Adding Graph in Excel Files....

    # Options...
    recent_op = 1
    month_select_op = 1
    view_graph_op = 1

    input_file = "req_trade.xlsx"
    #input_file = "req_trade_example.xlsx"

    read_data = read_req_excel_file(input_file)

    if recent_op == 1:
        result_list = crawling_recent_trade(read_data, month_select_op)
    else:
        result_list = crawling_all_trade(read_data, month_select_op)

    #print(result_list)

    write_excel_file(read_data, result_list, recent_op, view_graph_op)


# Main
if __name__ == "__main__":
    main()



