from selenium import webdriver
from selenium.webdriver.support.select import Select
import selenium as se
import time
import xlrd
import mysql.connector
from mysql.connector import Error
from mysql.connector import errorcode



def modemScrap():

    i = 1
    
    workbook = xlrd.open_workbook('ip2.xlsx')
    worksheet = workbook.sheet_by_index(0)

    print(worksheet.nrows)

    while i < worksheet.nrows:
    
        'hiding the browser,  for using the browser affichage remove all line and -> driver = webdriver.Chrome()'

        'Get http modem'
        
        ip = worksheet.cell(i,1).value
        print(ip)
        
        i+=1
        
     

        options = se.webdriver.ChromeOptions()
        options.add_argument("headless")

        driver = se.webdriver.Chrome(chrome_options=options)
        driver.get("http://"+ip)        
        login = driver.find_element_by_xpath("//*[@id=\"loginUsername\"]")
        login.send_keys("admin")
        mdp = driver.find_element_by_xpath("//*[@id=\"loginPassword\"]")
        mdp.send_keys("admin")
        GotoLogin = driver.find_element_by_xpath("/html/body/div[1]/div[2]/div[2]/div[2]/table/tbody/tr[2]/td/form/table/tbody/tr[2]/td/input")
        GotoLogin.submit()

   
        try:
                GoToWebInterfaceTC7210_Configuration= driver.find_element_by_xpath("/html/body/div[1]/form/center/a")
                GoToWebInterfaceTC7210_Configuration.click()
        except:
                GoToWebInterfaceTC7210_force = driver.find_element_by_xpath("/html/body/div[1]/div[2]/div[2]/div[2]/table/tbody/tr[2]/td/form/table/tbody/tr[3]/td/input[1]")
                GoToWebInterfaceTC7210_force.click()
                GoToWebInterfaceTC7210_Configuration_force = driver.find_element_by_class_name("button")
                GoToWebInterfaceTC7210_Configuration_force.click()

        VersionLogiciel_To_Scrap = driver.find_element_by_xpath('/html/body/div[1]/div[3]/div[2]/div[2]/table/tbody/tr[2]/td/form/table[1]/tbody/tr[4]/td[2]')
        CableModemIP_To_Scrap = driver.find_element_by_xpath('/html/body/div[1]/div[3]/div[2]/div[2]/table/tbody/tr[2]/td/form/table[2]/tbody/tr[4]/td[2]')


        CableModemIP_To_DB = CableModemIP_To_Scrap.text
        VersionLogiciel_To_DB =VersionLogiciel_To_Scrap.text

        print(VersionLogiciel_To_DB)
        print(CableModemIP_To_DB)

        try:
                'go to te wifi settings '
                GoToWifiParam = driver.find_element_by_xpath('/ html / body / div[1] / div[2] / ul / li[6] / a')
                GoToWifiParam.click()
                'scrap the 802.11n mode : auto or inactivr'

        except:
                'go to te wifi settings bridge '
                GoToWifiParamBridge = driver.find_element_by_xpath('/html/body/div[1]/div[2]/ul/li[2]/a')
                GoToWifiParamBridge.click()
        
        'scrap the 802.11n mode : auto or inactivr'


        Mode80211n_To_Scrap= Select(driver.find_element_by_xpath('/html/body/div[1]/div[3]/div[2]/div[2]/table/tbody/tr[2]/td/form/table/tbody/tr[5]/td[2]/select'))


        Mode80211n_Scraped= Mode80211n_To_Scrap.first_selected_option

        Mode80211n_FinalDB = Mode80211n_Scraped.text
        print(Mode80211n_FinalDB)

        BF_to_Scrap = Select(driver.find_element_by_xpath('/html/body/div[1]/div[3]/div[2]/div[2]/table/tbody/tr[2]/td/form/table/tbody/tr[6]/td[2]/select'))
        BF_Scraped = BF_to_Scrap.first_selected_option
        BF_FinalDB = BF_Scraped.text


        CanalWifiActuel_To_Scrap = driver.find_element_by_xpath('/html/body/div[1]/div[3]/div[2]/div[2]/table/tbody/tr[2]/td/form/table/tbody/tr[9]/td[2]')
        CanalWifiActuel_To_DB = CanalWifiActuel_To_Scrap.text

        Interface_To_Scrap = Select(driver.find_element_by_xpath('/html/body/div[1]/div[3]/div[2]/div[2]/table/tbody/tr[2]/td/form/table/tbody/tr[1]/td[2]/select'))

        Interface_Scraped = Interface_To_Scrap.first_selected_option
        Interface_Scraped_FinalDB = Interface_Scraped.text

        print(CanalWifiActuel_To_DB)
        print(BF_FinalDB)
        print(Interface_Scraped_FinalDB)

        
        GoToPrincipalNetwork = driver.find_element_by_xpath('/html/body/div[1]/div[3]/div[2]/div[1]/ul/li[3]/a')
        GoToPrincipalNetwork.click()

        Bandsteering_To_Scrap = Select(driver.find_element_by_xpath('/html/body/div[1]/div[3]/div[2]/div[2]/table/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/table[1]/tbody/tr[2]/td[2]/select'))

        Bandsteering_Scraped = Bandsteering_To_Scrap.first_selected_option
        Bandsteering_Scraped_FinalDB = Bandsteering_Scraped.text

        MainNetwork_To_Scrap = Select(driver.find_element_by_xpath('/html/body/div[1]/div[3]/div[2]/div[2]/table/tbody/tr[2]/td/form/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td[2]/select'))
        MainNetwork_Scraped = MainNetwork_To_Scrap.first_selected_option
        MainNetwork_Scraped_FinalDB = MainNetwork_Scraped.text

        print(Bandsteering_Scraped_FinalDB)
        print(MainNetwork_Scraped_FinalDB)

        try:
                GoToNetwork5G = driver.find_element_by_xpath('/html/body/div[1]/div[3]/div[2]/div[1]/ul/li[15]/a')
                GoToNetwork5G.click()

        except:
                GoToNetwork5G = driver.find_element_by_xpath('/html/body/div[1]/div[3]/div[2]/div[1]/ul/li[13]/a')
                GoToNetwork5G.click()

        Mode80211n5G_To_Scrap= Select(driver.find_element_by_xpath('/html/body/div[1]/div[3]/div[2]/div[2]/table/tbody/tr[2]/td/form/table/tbody/tr[4]/td[2]/select'))


        Mode80211n5G_Scraped= Mode80211n5G_To_Scrap.first_selected_option

        Mode80211n5G_FinalDB = Mode80211n5G_Scraped.text
        print(Mode80211n5G_FinalDB)

        BF5G_to_Scrap = Select(driver.find_element_by_xpath('/html/body/div[1]/div[3]/div[2]/div[2]/table/tbody/tr[2]/td/form/table/tbody/tr[5]/td[2]/select'))
        BF5G_Scraped = BF5G_to_Scrap.first_selected_option
        BF5G_FinalDB = BF5G_Scraped.text


        CanalWifiActuel5G_To_Scrap = driver.find_element_by_xpath('/html/body/div[1]/div[3]/div[2]/div[2]/table/tbody/tr[2]/td/form/table/tbody/tr[8]/td[2]')
        CanalWifiActuel5G_To_DB = CanalWifiActuel5G_To_Scrap.text

        Interface5G_To_Scrap = Select(driver.find_element_by_xpath('/html/body/div[1]/div[3]/div[2]/div[2]/table/tbody/tr[2]/td/form/table/tbody/tr[1]/td[2]/select'))

        Interface5G_Scraped = Interface5G_To_Scrap.first_selected_option
        Interface5G_Scraped_FinalDB = Interface5G_Scraped.text

        StatusDFS_To_Scrap = Select(driver.find_element_by_xpath('/html/body/div[1]/div[3]/div[2]/div[2]/table/tbody/tr[2]/td/form/table/tbody/tr[10]/td[2]/select'))
        StatusDFS_To_Scraped = StatusDFS_To_Scrap.first_selected_option
        StatusDFS_To_Scraped_FinalDB = StatusDFS_To_Scraped.text

        print(CanalWifiActuel5G_To_DB)
        print(BF5G_FinalDB)
        print(Interface5G_Scraped_FinalDB)
        print(StatusDFS_To_Scraped_FinalDB)

                    
        connection = mysql.connector.connect(host='',
        database='infoclient',
        user='',
        password=''!')
        cursor = connection.cursor(prepared=True)
        sql_insert_query = """INSERT INTO modem
        (version, cmip, mode, canal, bande, bandsteering, interface,mode_5G, canal_5G, bande_5G, interface_5G, DFS) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"""
        insert_tuple = (VersionLogiciel_To_DB,CableModemIP_To_DB, Mode80211n_FinalDB, CanalWifiActuel_To_DB, BF_FinalDB, Bandsteering_Scraped_FinalDB, Interface_Scraped_FinalDB, Mode80211n5G_FinalDB,CanalWifiActuel5G_To_DB,BF5G_FinalDB,Interface5G_Scraped_FinalDB,StatusDFS_To_Scraped_FinalDB)
        cursor.execute(sql_insert_query, insert_tuple)
        connection.commit()
        print ("Record inserted successfully into python_users table")
            
        if(connection.is_connected()):
                cursor.close()
                connection.close()
                print("MySQL connection is closed")

    return
    
modemScrap()
