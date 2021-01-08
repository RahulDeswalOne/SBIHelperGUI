from selenium import webdriver
import xlwings as xw
# import traceback
# import sys
import selenium.webdriver.support.ui

def login_sbi(username, usrpass, brw):
    global browser
    if brw == 1:
        browser = webdriver.Edge('../webdriver/msedgedriver')
    else:   # brw == 2:
        browser = webdriver.Chrome('../webdriver/chromedriver')

    browser.get("https://yonobusiness.sbi/login/yonobusinesslogin")
    browser.find_element_by_xpath('//*[@id="userName"]').send_keys(username)
    browser.find_element_by_xpath('//*[@id="password"]').send_keys(usrpass)
    browser.find_element_by_xpath('//*[@id="Capchalogin"]').click()
    browser.implicitly_wait(20)
    browser.find_element_by_xpath(
        '//*[@id="postlogin-sbi"]/div[1]/div[1]/app-side-nav-bar/div[2]/aside/ul/li[2]/a').click()


def exit_sbi():
    browser.find_element_by_xpath('//*[@id="useful-links"]/div/ul/li[6]/a').click()

def pending_ent():
    wb = xw.books("Online Cheque.xlsm")
    ws = wb.sheets['current']
    nRws = int(ws.range("b1").value)
    print("Pending Cheques Module", nRws)
    wb = xw.books("Online Cheque.xlsm")
    ws = wb.sheets['Current']
    nRws = int(ws.range('k1').value)

    # Removing old data from Excel Sheet
    ws.range((4, 11), (nRws + 4, 18)).value = ""

    browser.find_element_by_link_text('Reports').click()
    browser.find_element_by_link_text('Pending Transactions').click()
    browser.find_element_by_id('startDate').click()
    browser.find_element_by_xpath('/html/body/div[3]/div[1]/table/thead/tr[2]/th[1]').click()
    trows = browser.find_element_by_class_name('datepicker-days')
    trows.find_element_by_xpath('/html/body/div[3]/div[1]/table/tbody/tr[1]/td[1]').click()
    browser.find_element_by_xpath(
        '/html/body/div[1]/div[1]/div/div[2]/div/div[1]/form/div/div[3]/input[1]').click()
    browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[2]/div/div/div[2]/div/div')
    rows = len(browser.find_elements_by_xpath(
        '//*[@id="contentDiv"]/div/div/div[2]/div/div/div[3]/table/tbody/tr'))
    cols = len(browser.find_elements_by_xpath(
        '//*[@id="contentDiv"]/div/div/div[2]/div/div/div[3]/table/thead/tr/th'))

    for rs in range(1, rows + 1):
        for cl in range(1, cols + 1):
            data = browser.find_element_by_xpath(
                '//*[@id="contentDiv"]/div/div/div[2]/div/div/div[3]/table/tbody/tr[' + str(
                    rs) + ']/td[' + str(cl) + ']')
            data = data.text
            ws.range(rs + 3, cl + 10).value = data

    # for changing date
    for rs in range(1, nRws-1):
        chDate = str(ws.range(rs + 3, 14).value)
        left = chDate[0:2]
        left = int(left)
        nRight = chDate[8:10]
        nLen = len(chDate)
        if int(left) > 12 and nLen < 11:
            nYr = chDate[6:10]
            nMn = chDate[3:5]
            nDt = chDate[0:2]
            oUtput = '|' + nDt + '-' + nMn + '-' + nYr
            fDate = oUtput[1:11]
            ws.range(rs + 3, 14).value = fDate
        elif int(nRight) < 13:
            nYr = chDate[0:4]
            nMn = chDate[5:7]
            nDt = chDate[8:10]
            oUtput = '|' + nDt + '-' + nMn + '-' + nYr
            fDate = oUtput[1:11]
            ws.range(rs + 3, 14).value = fDate
    print("Done")


def statement_SBI():
    print("Statement Module")
    wbSt = xw.books("Tally Import Utility V_6.1.1.xlsm")
    wsSt = wbSt.sheets['bnkimp']

    # Removing old data from Excel Sheet
    nRws = int(wsSt.range('l1').value)
    wsSt.range((2, 1), (nRws + 1, 12)).value = ""

    # Going to Statement page
    browser.find_element_by_xpath('//*[@id="navbar"]/div[1]/a[1]').click()
    browser.find_element_by_xpath('//*[@id="contentDiv"]/div/div[2]/div[2]/div[2]/h3/a').click()
    browser.find_element_by_xpath('//*[@id="fromDate"]').click()
    browser.find_element_by_xpath('/html/body/div[3]/div[1]/table/thead/tr[2]/th[1]').click()
    tHead = browser.find_element_by_xpath('/html/body/div[3]/div[1]/table/tbody')
    print(tHead.text)
    i = 0
    while i == 31:
        print("Step1")
        datePicker = tHead.text
        print(datePicker)
        print('Printed')

    browser.find_element_by_xpath('//*[@id="toDate"]').click()
    browser.find_element_by_xpath('//*[@id="view"]').click()
    browser.find_element_by_xpath('//*[@id="contentDiv"]/form/div/div[2]/div[4]/ul/li/a').click()
    rStatus = browser.find_element_by_id("accountNo").is_selected()
    print(rStatus)
    browser.implicitly_wait(200)
    browser.find_element_by_xpath('//*[@id="demo3"]/div/div/div[3]/table/tbody/tr/td[1]/label')
    rStatus = browser.find_element_by_id("accountNo").is_selected()
    print(rStatus)
    browser.find_element_by_xpath('//*[@id="contentDiv"]/form/div/div[2]/div[5]/input[1]').click()
    browser.find_elements_by_xpath(
        '//*[@id="contentDiv"]/div[1]/div[2]/div[6]/div[6]/div[3]/table')
    rows = len(browser.find_elements_by_xpath(
        '//*[@id="contentDiv"]/div[1]/div[2]/div[6]/div[6]/div[3]/table/thead/tr'))
    cols = len(browser.find_elements_by_xpath(
        '//*[@id="contentDiv"]/div[1]/div[2]/div[6]/div[6]/div[3]/table/thead/tr/th'))

    for rs in range(1, rows + 1):
        for cl in range(1, cols + 1):
            data = browser.find_element_by_xpath(
                '//*[@id="contentDiv"]/div[1]/div[2]/div[6]/div[6]/div[3]/table/thead/tr[' + str(
                    rs) + ']/td[' + str(cl) + ']')
            data = data.text
            wsSt.range(rs + 3, cl + 10).value = data


def prepare_chq(trn_pwd):
    wb = xw.books("Online Cheque.xlsm")
    ws = wb.sheets['current']
    ws1 = wb.sheets['xpath']
    cls = ws.range("b4:h44").columns.count
    trnPass = ws1.range("aa3").value

    nRws = int(ws.range("b1").value)
    print("Payment Module", nRws)
    for rw in range(4, nRws + 4):
        print('first for')
        bankType = ws.range(rw, 1).value
        chqNum = ws.range(rw, 2).value
        opartyNam = str(ws.range(rw, 3).value)
        chqAmt = int(ws.range(rw, 4).value)
        oRemark = str(ws.range(rw, 5).value)
        remarkCd = int(ws.range(rw, 8).value)
        print(chqNum, chqAmt, oRemark, remarkCd)
        if oRemark is None:
            print('first if')
            nRemark = 'On Account'
        else:
            nRemark = oRemark
            nRemark = nRemark[0:50]
        partyNam = opartyNam[0:50]
        if chqNum is None and partyNam is not None and bankType == 'OtherBank' and chqAmt > 100:
            print('other payment')
            browser.find_element_by_xpath('//*[@id="navbar"]/div[1]/a[1]').click()
            payment = browser.find_element_by_xpath('//*[@id="navbar"]/div[2]/a[1]')
            payment.click()
            othBank = browser.find_element_by_xpath('//*[@id="contentDiv"]/div/div[2]/div[3]/div[2]/h3/a')
            browser.find_element_by_xpath('//*[@id="contentDiv"]/div/div[2]/div[2]/div[2]/h3/a')
            browser.find_element_by_xpath('//*[@id="contentDiv"]/div/div[4]/div[3]/div[2]/h3/a')
            othBank.click()
            rTGS = browser.find_element_by_xpath('//*[@id="iconRTGS"]')
            nEFT = browser.find_element_by_xpath('//*[@id="iconNEFT"]')
            if chqAmt > 199999:
                print('rtgs module')
                rTGS.click()
                browser.find_element_by_xpath('//*[@id="debitAmount"]').send_keys(chqAmt)
                purposeCode = browser.find_element_by_xpath('//*[@id="purposeCode"]')
                srchBene = browser.find_element_by_xpath('//*[@id="otherTd"]/a')
                srchBRB = browser.find_element_by_xpath('//*[@id="container"]/div[1]/label[3]/input')
                browser.find_element_by_xpath('//*[@id="remarks"]').send_keys(nRemark)
                drp1 = selenium.webdriver.support.ui.Select(purposeCode)
                drp1.select_by_value('CORT:Trade settlement Payment')
                srchBene.click()
                srchBRB.click()
                browser.find_element_by_id('srchInput').send_keys(partyNam)
                browser.find_element_by_id('srchBtn').click()
                browser.find_element_by_id('srdr0').click()
                browser.find_element_by_name('acceptTerms').click()
                browser.find_element_by_id('confirmSubmit').click()
                browser.find_element_by_id('transactionPwd').send_keys(trnPass)
                browser.find_element_by_id('confirmSubmit').click()
                chNmbr = browser.find_element_by_xpath(
                    '//*[@id="contentDiv"]/form/div[1]/div[2]/div[3]/div[9]').text
                ws.range(rw, 2).value = chNmbr[1:11]
                print('RTGS Cheque done! for ', partyNam, chqAmt)
            else:
                def purpsel(i):
                    switcher = {
                        0: 'Other Payments',
                        1: 'Salary Payment',
                        2: 'Payment towards Invoice Or Bill',
                    }
                    return switcher.get(i, "Invalid Selection Value Please enter 0, 1 or 2")

                trasPrpsSel = purpsel(remarkCd)
                nEFT.click()
                browser.find_element_by_xpath('//*[@id="debitAmount"]').send_keys(chqAmt)
                browser.find_element_by_xpath('//*[@id="creditAmount"]').send_keys(chqAmt)
                transPurpose = browser.find_element_by_xpath('//*[@id="purposeoftransact"]')
                srchBene = browser.find_element_by_xpath('//*[@id="otherTd"]/a')
                srchBRB = browser.find_element_by_xpath('//*[@id="container"]/div[1]/label[3]/input')
                drp2 = selenium.webdriver.support.ui.Select(transPurpose)
                drp2.select_by_value(trasPrpsSel)
                if trasPrpsSel == "Other Payments":
                    browser.find_element_by_xpath('//*[@id="remarks"]').send_keys(nRemark)
                srchBene.click()
                srchBRB.click()
                browser.find_element_by_id('srchInput').send_keys(partyNam)
                browser.find_element_by_id('srchBtn').click()
                browser.find_element_by_id('srdr0').click()
                browser.find_element_by_name('acceptTerms').click()
                browser.find_element_by_id('confirmSubmit').click()
                browser.find_element_by_id('transactionPwd').send_keys(trnPass)
                # browser.find_element_by_name('acceptTerms').click()
                browser.find_element_by_id('confirmSubmit').click()
                chNmbr = browser.find_element_by_xpath(
                    '//*[@id="contentDiv"]/form/div[1]/div[2]/div[3]/div[9]').text
                ws.range(rw, 2).value = chNmbr[1:11]
                print('NEFT Cheque done! for ', opartyNam, chqAmt)
        elif chqNum is None and partyNam is not None and bankType == 'SameBank':
            print('same bank module')
            browser.find_element_by_xpath('//*[@id="navbar"]/div[1]/a[1]').click()
            payment = browser.find_element_by_xpath('//*[@id="navbar"]/div[2]/a[1]')
            payment.click()
            browser.find_element_by_xpath('//*[@id="contentDiv"]/div/div[2]/div[3]/div[2]/h3/a')
            sameBank = browser.find_element_by_xpath('//*[@id="contentDiv"]/div/div[2]/div[2]/div[2]/h3/a')
            sameBank.click()
            browser.implicitly_wait(20)
            browser.find_element_by_xpath('//*[@id="clx"]/img').click()
            browser.find_element_by_xpath('//*[@id="debitAmount"]').send_keys(chqAmt)
            browser.find_element_by_xpath('//*[@id="creditAmount"]').send_keys(chqAmt)
            browser.find_element_by_xpath(
                '//*[@id="fundstransfer"]/div[1]/div[2]/div[8]/div/input').send_keys(nRemark)
            srchBene = browser.find_element_by_xpath('//*[@id="otherTd"]/a')
            srchBRB = browser.find_element_by_xpath('//*[@id="container"]/div[1]/label[3]/input')
            srchBene.click()
            srchBRB.click()
            browser.find_element_by_id('srchInput').send_keys(partyNam)
            browser.find_element_by_id('srchBtn').click()
            browser.find_element_by_id('srdr0').click()
            browser.find_element_by_xpath('//*[@id="divDefault"]/div[4]/input[2]').click()
            browser.find_element_by_id('transactionPwd').send_keys(trnPass)
            browser.find_element_by_xpath('//*[@id="divDefault"]/div[4]/input[2]').click()
            chNmbr = browser.find_element_by_xpath(
                '//*[@id="contentDiv"]/form/div[1]/div[2]/div[3]/div[9]').text
            ws.range(rw, 2).value = chNmbr[1:11]
            print('Transfer Cheque done! for ', partyNam, chqAmt)
        elif chqAmt == 100:
            print('imps module')
            browser.find_element_by_xpath('//*[@id="navbar"]/div[1]/a[1]').click()
            payment = browser.find_element_by_xpath('//*[@id="navbar"]/div[2]/a[1]')
            payment.click()
            iMPS = browser.find_element_by_xpath('//*[@id="contentDiv"]/div/div[4]/div[3]/div[2]/h3/a')
            iMPS.click()
            browser.find_element_by_xpath('//*[@id="transacType"]').click()
            browser.find_element_by_xpath('//*[@id="confirmSubmit"]').click()
            browser.switch_to.alert.accept()
            browser.find_element_by_xpath('//*[@id="debitAmount"]').send_keys(chqAmt)
            browser.find_element_by_xpath('//*[@id="creditAmount"]').send_keys(chqAmt)
            browser.find_element_by_xpath('//*[@id="debitRemarks"]').send_keys(nRemark)
            browser.find_element_by_xpath('//*[@id="otherTd"]/a').click()
            browser.find_element_by_xpath('//*[@id="container"]/div[1]/label[3]/input').click()
            browser.find_element_by_id('srchInput').send_keys(partyNam)
            browser.find_element_by_id('srchBtn').click()
            browser.find_element_by_id('srdr0').click()
            browser.find_element_by_name('acceptTerms').click()
            browser.find_element_by_id('confirmSubmit').click()
            browser.find_element_by_id('transactionPwd').send_keys(trnPass)
            # browser.find_element_by_name('acceptTerms').click()
            browser.find_element_by_id('confirmSubmit').click()
            chNmbr = browser.find_element_by_xpath(
                '//*[@id="myTabContent"]/form/div/div/div[3]/div[9]').text
            ws.range(rw, 2).value = chNmbr[1:11]
            print('iIMPS Cheque done! for ', opartyNam, chqAmt)