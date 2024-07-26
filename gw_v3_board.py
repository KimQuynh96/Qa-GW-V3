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
    commons.Title("Folders")
    
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
            j = 1
            Name_Edit = Folder + "[Edited]"
            while j <= List_Name :
                Edit = List_Folder.find_element_by_xpath(data["board"]["folder_checkbox"] % str(j))
                Name = Edit.text
                
                if Name == Folder:
                    Edit.click()
                    print("oki")
                    Folder_Edit = driver.find_element_by_xpath(data["board"]["folder_name"])
                    Folder_Edit.send_keys(Keys.CONTROL + "a")
                    Folder_Edit.send_keys(Keys.DELETE)
                    Folder_Edit.send_keys(Keys.RETURN)
                    Folder_Edit.send_keys(Name_Edit)
                    print('Name_Edit :',Name_Edit)
                    break
                j += 1
        
        
        
        
        
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