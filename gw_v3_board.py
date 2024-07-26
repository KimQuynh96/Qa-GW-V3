from gw_v3_set_up import driver,data,Keys,EC,By,WebDriverWait,json,File,os,Param,chrome_options,pyautogui,GetSystemMetrics
from gw_v3_functions import commons, time


def AccessBoard() :
    commons.Title("MENU BOARD")
    commons.ClickElementWithXpath(data["board"]["menu"])
    
    if commons.IsDisplayedByXpath(data["board"]["write"]) == True :
        Access = True
        commons.WriteOnExcel(data["board_excel"]["menu"]["pass"])
    else :
        Access = False
        commons.WriteOnExcel(data["board_excel"]["menu"]["fail"])
    return Access

def Folder():

    # Add Folder #
    commons.Title("Folders : Add")
    
    Folder = str(commons.Time())
    Folder = Folder[5: 16]+"]"
    time.sleep(1)
    driver.find_element_by_xpath(data["board"]["folder_name"]).send_keys(Folder)
    commons.Content("Input folder name")
    
    driver.find_element_by_xpath(data["board"]["folder_save"]).click()
    commons.Content("Click on button save")
        
    List_Folder = driver.find_element_by_css_selector(data["board"]["folder_list"])    
    time.sleep(3)
    List_Name = List_Folder.find_elements_by_class_name(data["board"]["folder_input"])   
    List_Name = commons.TotalData(List_Name)
    
    if  List_Name != 0 :
        i = 1
        Create = False
        while i <= List_Name :
            Name = List_Folder.find_element_by_xpath(data["board"]["folder_checkbox"] % str(i)).text
            if Name == Folder:
                Create = True
                commons.WriteOnExcel(data["board_excel"]["setting"]["create"]["pass"])
                break
            i += 1
        
        if  Create == False :
            commons.WriteOnExcel(data["board_excel"]["setting"]["create"]["fail"])
        
        # Edit Folder #
        if Create == True :
            commons.Title("Folders : Edit")
            j = 1
            Name_Edit = Folder + "[Edited]"
            while j <= List_Name :
                Edit = List_Folder.find_element_by_xpath(data["board"]["folder_checkbox"] % str(j))
                Name = Edit.text
                
                if Name == Folder:
                    Edit.click()
                    Folder_Edit = driver.find_element_by_xpath(data["board"]["folder_name"])
                    time.sleep(1)
                    Folder_Edit.send_keys(Keys.CONTROL + "a")
                    Folder_Edit.send_keys(Keys.DELETE)
                    Folder_Edit.send_keys(Keys.RETURN)
                    time.sleep(1)
                    Folder_Edit.send_keys(Name_Edit)
                    commons.Content("Input name edit")
                    driver.find_element_by_xpath(data["board"]["folder_save"]).click()
                    commons.Content("Click on button save edit")
                    
                    m = 1 
                    Edited = False
                    time.sleep(2)
                    while m <= List_Name :
                        After_Edit = List_Folder.find_element_by_xpath(data["board"]["folder_checkbox"] % str(m))
                        After_Name = After_Edit.text

                        if After_Name == Name_Edit:
                            Edited = True
                            commons.WriteOnExcel(data["board_excel"]["setting"]["edit"]["pass"])
                            break
                        m += 1
                    if  Edited == False :
                        commons.WriteOnExcel(data["board_excel"]["setting"]["edit"]["fail"])
                        
                    break
                j += 1
        
        '''
        # Share Folder #
        if Create == True :
            commons.Title("Folders : Share")
            j = 1
            while j <= List_Name :
                Find = List_Folder.find_element_by_xpath(data["board"]["folder_checkbox"] % str(j))
                Name = Find.text
                Find_Name_1 = (Name.rfind(Folder))
                Find_Name_2 = (Name.rfind(Name_Edit))
                if Find_Name_1 != -1 or Find_Name_2 != -1 :
                    Find.click()
                    driver.find_element_by_css_selector(data["board"]["folder_share"]).click()
                    commons.Content("Click on button shared")
                    driver.find_element_by_css_selector(data["board"]["folder_invite"]).click()
                    commons.Content("Click on button invite")
                    
                    time.sleep(1)
                    Global = driver.find_element_by_class_name("css-qa6tqs")
                    Global.find_element_by_xpath(data["board"]["folder_open"]).click()
                j += 1
        '''          
        # Delete Folder #
        if Create == True :
            commons.Title("Folders : Delete")
            driver.find_element_by_xpath(data["board"]["folder_delete"]).click()
            if commons.IsDisplayedByXpath(data["board"]["folder_close"]) == True :
                commons.Content("Click on button delete")
                driver.find_element_by_xpath(data["board"]["folder_ok"]).click()
                if  commons.IsDisplayedByXpath(data["board"]["folder"]) == False :
                    commons.Content("Click on button OK")
                    
                    time.sleep(3)
                    List_After = List_Folder.find_elements_by_class_name(data["board"]["folder_input"])   
                    List_After = commons.TotalData(List_After)
                    if  List_Name == List_After + 1:
                        commons.WriteOnExcel(data["board_excel"]["setting"]["delete"]["pass"])
                    
                    else :
                        commons.WriteOnExcel(data["board_excel"]["setting"]["delete"]["fail"])
                    
                else :
                    commons.WriteOnExcel(data["board_excel"]["setting"]["delete"]["fail"])
            else :
                commons.WriteOnExcel(data["board_excel"]["setting"]["delete"]["fail"])
    else :
        commons.WriteOnExcel(data["board_excel"]["setting"]["create"]["fail"])

def Settings():
    commons.Title("Settings")
    commons.ClickElementWithText(data["board"]["setting"])
    if commons.IsDisplayedByXpath(data["board"]["folder_btn"]) == True :
        Settings = True
        commons.WriteOnExcel(data["board_excel"]["setting"]["access"]["pass"])
    else :
        Settings = False
        commons.WriteOnExcel(data["board_excel"]["setting"]["access"]["fail"])
        
    if  Settings == True :
        '''
        try :
            Folder()
        except :
            commons.CaseFail("Error Xpath : Folder")
        '''
        Folder()
def MenuBoard():
    Access = AccessBoard() 
    if  Access == True :
        Settings()