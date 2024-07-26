from gw_v3_set_up import driver,data,Keys,EC,By,WebDriverWait,json,File,os,Param,chrome_options,pyautogui,GetSystemMetrics
from gw_v3_functions import commons, time


def AccessMail():
    commons.Title("MENU MAIL")
    commons.ClickElementWithXpath(data["mail"]["menu"])
    # Frame 1
    driver.switch_to.frame(commons.FindElementById("newMailIframe"))
    if commons.IsDisplayedByXpath(data["mail"]["compose"]) == True :
        Access = True
        commons.WriteOnExcel(data["mail_excel"]["menu"]["pass"])
    else :
        Access = False
        commons.WriteOnExcel(data["mail_excel"]["menu"]["fail"])
    # Close Frame 1
    commons.SwitchToDefaultContent()
    
    return Access

def WritingSetting():
    commons.Title("Writing Setting")
    Reply_To = driver.find_element_by_xpath(data["mail"]["reply_to"])
    Reply_To.clear()
    Reply_To.send_keys(data["mail"]["mail_to"])
    commons.Content("Input Reply-To")
    
    driver.find_element_by_xpath(data["mail"]["setting_save"]).click()
    commons.Content("Click on button save")
    commons.WriteOnExcel(data["mail_excel"]["setting"]["writing"]["pass"])
    
    
def Signature():
    commons.Title("Signature")
    commons.ClickElementWithXpath(data["mail"]["signature"])
    if  commons.IsDisplayedByXpath(data["mail"]["sig_btn"]) == True :
        Access = True 
        commons.WriteOnExcel(data["mail_excel"]["setting"]["signature"]["access"]["pass"])
    else :
        Access = False
        commons.WriteOnExcel(data["mail_excel"]["setting"]["signature"]["access"]["fail"])    
           
    if  Access == True :
        commons.Title("Signature  : Add")
        Total_Before = driver.find_elements_by_xpath(data["mail"]["sig_input"])
        Total_Before = len(Total_Before) 
        
        List_Signature = []
        i = 1
        while i <= Total_Before:
            Name = driver.find_element_by_xpath(data["mail"]["sig_name"] % str(i)).text
            List_Signature.append(Name)
            i += 1
        
        driver.find_element_by_xpath(data["mail"]["sig_btn"]).click()
        
        if  commons.IsDisplayedByCss(data["mail"]["sig_attach"]) == True :
            Add = True 
            commons.Content("Click on button add signature")
        else :
            Add = False
            commons.CaseFail("Click on button add signature")

        
        if  Add == True :
            driver.find_element_by_xpath(data["mail"]["sig_text"]).click()
            commons.Content("Click on radio text")
            commons.SwitchToFrameEditor(data["editor_frame"])
            commons.Wait10s_InputElement(data["editor_input"],data["sig_text"])
            commons.SwitchToDefaultContent
            commons.Content("Input signature is text")
        
            # Close Frame 2
            driver.switch_to.default_content()
            
            # Frame 3
            driver.switch_to.frame(commons.FindElementById("newMailIframe"))
            driver.find_element_by_xpath(data["mail"]["sig_save"]).click()
            
            if  commons.IsDisplayedByXpath(data["mail"]["sig_img"]) == False :
                Save = True
                commons.Content("Click on button save")
            else :
                Save = False
                commons.CaseFail("Click on button save")
            
            if Save == True :
                Total_After = driver.find_elements_by_xpath(data["mail"]["sig_input"])
                Total_After = len(Total_After) 
                if  Total_After == Total_Before + 1 :
                    commons.WriteOnExcel(data["mail_excel"]["setting"]["signature"]["create"]["pass"])
                else:
                    commons.WriteOnExcel(data["mail_excel"]["setting"]["signature"]["create"]["fail"])   
                    
            # Delete Signature 
            if Save == True :
                time.sleep(1)
                Total_Delete = driver.find_elements_by_xpath(data["mail"]["sig_input"])
                Total_Delete = len(Total_Delete) 
            
                j = 1
                while j <= Total_Delete:
                    Name = driver.find_element_by_xpath(data["mail"]["sig_name"] % str(j)).text
                    if Name in List_Signature :
                        j += 1
                    else :
                        driver.find_element_by_xpath(data["mail"]["sig_checkbox"] % str(j)).click()
                        commons.Content("Click on checkbox")
                        driver.find_element_by_xpath(data["mail"]["sig_delete"]).click()
                        commons.Content("Click on button delete")
                        
                        time.sleep(1)
                        Total_After = driver.find_elements_by_xpath(data["mail"]["sig_input"])
                        Total_After = len(Total_After) 
                        if  Total_Delete == Total_After + 1 :
                            commons.WriteOnExcel(data["mail_excel"]["setting"]["signature"]["delete"]["pass"])
                        else:
                            commons.WriteOnExcel(data["mail_excel"]["setting"]["signature"]["delete"]["fail"])  
                        break    
        else :
            commons.WriteOnExcel(data["mail_excel"]["setting"]["signature"]["create"]["fail"])    
            
def AutoSort():
    commons.Title("Auto Sort")
    commons.ClickElementWithXpath(data["mail"]["auto_sort"])
    if  commons.IsDisplayedByCss(data["mail"]["auto_btn"]) == True :
        Access = True 
        commons.WriteOnExcel(data["mail_excel"]["setting"]["auto"]["access"]["pass"])
    else :
        Access = False
        commons.WriteOnExcel(data["mail_excel"]["setting"]["auto"]["access"]["fail"])        
    
    if  Access == True :
        commons.Title("Auto Sort : Add")
        Table = driver.find_element_by_xpath(data["mail"]["auto_table"])
        Total_Before =  Table.find_elements_by_tag_name("tr")
        Total_Before =  commons.TotalData(Total_Before) + 1
        
        driver.find_element_by_css_selector(data["mail"]["auto_btn"]).click()
        commons.Content("Click on button add")
        time.sleep(2)
        driver.find_element_by_id(data["mail"]["auto_from"]).send_keys(data["mail"]["auto_from_email"])
        commons.Content("Input from")
        driver.find_element_by_xpath(data["mail"]["auto_to"]).send_keys(data["mail"]["auto_to_email"])
        commons.Content("Input to")
        driver.find_element_by_xpath(data["mail"]["auto_subject"]).send_keys(str(commons.Time()))
        commons.Content("Input subject")
        driver.find_element_by_xpath(data["mail"]["auto_save"]).click()
        commons.Content("Click on button save")
        
        time.sleep(2)
        Table = driver.find_element_by_xpath(data["mail"]["auto_table"])
        Total_After =  Table.find_elements_by_tag_name("tr")
        Total_After =  commons.TotalData(Total_After) + 1
        if  Total_After == Total_Before + 1 :
            commons.WriteOnExcel(data["mail_excel"]["setting"]["auto"]["create"]["pass"])
        else :
            commons.WriteOnExcel(data["mail_excel"]["setting"]["auto"]["create"]["fail"])
 
def VacationAutoReplies():
    commons.Title("Vacation Auto-Replies")
    commons.ClickElementWithXpath(data["mail"]["vaca_auto"])
    if  commons.IsDisplayedById(data["mail"]["vaca_mess"]) == True :
        Access = True 
        commons.WriteOnExcel(data["mail_excel"]["setting"]["vacation"]["access"]["pass"])
    else :
        Access = False
        commons.WriteOnExcel(data["mail_excel"]["setting"]["vacation"]["access"]["fail"])   
    
    
    if  Access == True :
        commons.Title("Vacation Auto-Replies : Add")
        if commons.IsDisplayedByCss(data["mail"]["vaca_on"]) == True :
            driver.find_element_by_css_selector(data["mail"]["vaca_on"]).click()
            commons.Content("Click turn on radio")
        Message = driver.find_element_by_id(data["mail"]["vaca_mess"])
        Message.clear()
        Message.send_keys(str(commons.Time()))
        commons.Content("Input Message")
        driver.find_element_by_xpath(data["mail"]["vaca_save"]).click()
        commons.WriteOnExcel(data["mail_excel"]["setting"]["vacation"]["create"]["pass"])

def BlockedAddresses():
    commons.Title("Blocked Addresses")
    commons.ClickElementWithXpath(data["mail"]["block"])
    if  commons.IsDisplayedById(data["mail"]["block_input"]) == True :
        Access = True 
        commons.WriteOnExcel(data["mail_excel"]["setting"]["block"]["access"]["pass"])
    else :
        Access = False
        commons.WriteOnExcel(data["mail_excel"]["setting"]["block"]["access"]["fail"])
    
    if  Access == True :
        
        No_Data = commons.IsDisplayedByXpath(data["mail"]["block_data"])
        if  No_Data == False :  
            i = 1    
            Table = driver.find_element_by_xpath(data["mail"]["auto_table"])
            Total_Before =  Table.find_elements_by_tag_name("tr")
            Total_Before =  commons.TotalData(Total_Before) 
            while  i < Total_Before + 1 :
                Email = driver.find_element_by_xpath(data["mail"]["block_email"] % str(i)).text
                if  Email == data["mail"]["block_text"] :
                    commons.Title("Blocked Addresses : Delete")
                    Checkbox = driver.find_element_by_xpath(data["mail"]["block_checkbox"] % str(i))
                    Checkbox.click()
                    commons.Content("Click on checkbox")
                    time.sleep(1)
                    driver.find_element_by_xpath(data["mail"]["block_delete"]).click()
                    commons.Content("Click on button delete")
                    break
                i += 1
        
        
        commons.Title("Blocked Addresses : Add")
        time.sleep(1)
        driver.find_element_by_id(data["mail"]["block_input"]).send_keys(data["mail"]["block_text"])
        commons.Content("Input email address")
        driver.find_element_by_xpath(data["mail"]["block_add"]).click()
        commons.Content("Click on button add ")
        
        
        No_Data = commons.IsDisplayedByXpath(data["mail"]["block_data"])
        if  No_Data == False :       
            i = 1  
            Create = False  
            Table  = driver.find_element_by_xpath(data["mail"]["auto_table"])
            Total_After =  Table.find_elements_by_tag_name("tr")
            Total_After =  commons.TotalData(Total_After) 
            while  i < Total_After + 1 :
                Email = driver.find_element_by_xpath(data["mail"]["block_email"] % str(i)).text
                if  Email == data["mail"]["block_text"] :
                    Create = True
                    commons.WriteOnExcel(data["mail_excel"]["setting"]["block"]["create"]["pass"])
                    break
                i += 1
            if  Create == False :
                commons.WriteOnExcel(data["mail_excel"]["setting"]["block"]["create"]["fail"])
        else :
            commons.WriteOnExcel(data["mail_excel"]["setting"]["block"]["create"]["fail"])
           
        
def Folders():
    
    commons.Title("Folders")
    commons.ClickElementWithXpath(data["mail"]["folder"])
    if  commons.IsDisplayedByXpath(data["mail"]["folder_box"]) == True :
        Access = True 
        commons.WriteOnExcel(data["mail_excel"]["setting"]["folder"]["access"]["pass"])
    else :
        Access = False
        commons.WriteOnExcel(data["mail_excel"]["setting"]["folder"]["access"]["fail"])
    
    if  Access == True :
        
        # Button Add Folder #
        commons.Title("Folders : Add")
        commons.ClickElementWithXpath(data["mail"]["folder_add"])
        if commons.IsDisplayedByXpath(data["mail"]["folder_reset"]) == True :
            Add = True 
            commons.Content("Click on button add folder")
        else :
            Add = False
            commons.CaseFail("Click on button add folder")
            commons.WriteOnExcel(data["mail_excel"]["setting"]["folder"]["create"]["fail"])
        
        # Add Folder #
        if  Add == True :
            Folder = "Folder : " + str(commons.Time())
            Folder_Name = driver.find_element_by_css_selector(data["mail"]["folder_input"])
            Folder_Name.send_keys(Folder)
            commons.Content("Input folder name")
            driver.find_element_by_xpath(data["mail"]["folder_save"]).click()
            commons.Content("Click on button save folder")
            
            No_Data = commons.IsDisplayedByXpath(data["mail"]["folder_no"])
            if  No_Data == False :
                Create = False
                List_Folder =  driver.find_elements_by_xpath(data["mail"]["folder_item"])
                List_Folder =  commons.TotalData(List_Folder)
                List_Folder = List_Folder /2
    
                i = 1
                while i <= List_Folder :
                    Name = driver.find_element_by_xpath(data["mail"]["folder_name"] % str(i)).text
                    Name = Name[int(Name.rfind("F")): int(Name.rfind("("))]
                    if  Name == Folder :
                        Create = True
                        commons.WriteOnExcel(data["mail_excel"]["setting"]["folder"]["create"]["pass"])
                        break
                    i += 1
                
                
                # Check folder display left folder #    
                commons.Title("Folders : Show left menu")   
                if Create == True : 
                    if commons.IsDisplayedByCss(data["mail"]["folder_search"]) == True :
                        Display = False
                        Folders_Left = driver.find_element_by_xpath(data["mail"]["folder_left"])
                        Folders_Left_Name  = Folders_Left.find_elements_by_tag_name("li")
                        Total_Folder = commons.TotalData(Folders_Left_Name)
                        i = 1
                        while i <= Total_Folder :
                            Left_Name = Folders_Left.find_element_by_xpath(data["mail"]["folder_left_name"] %str(i)).text
                            if  Left_Name == Folder :
                                Display = True
                                commons.WriteOnExcel(data["mail_excel"]["setting"]["folder"]["display"]["pass"])
                                break
                            i += 1
                        if  Display == False :
                            commons.WriteOnExcel(data["mail_excel"]["setting"]["folder"]["display"]["fail"])
                        
                
                if Create == True : 
                    
                    #   Backup 
                    commons.Title("Folders : Backup")
                    while i <= List_Folder :
                        Name_Text = driver.find_element_by_xpath(data["mail"]["folder_name"] % str(i))
                        Name = Name_Text.text
                        Name = Name[int(Name.rfind("F")): int(Name.rfind("("))]
                        if  Name == Folder :
                            Name_Text.click()
                            commons.Content("Click on folder name")
                            driver.find_element_by_xpath(data["mail"]["folder_backup"]).click()
                            Name = Name.replace(":","_")
                            time.sleep(3)
                            Download = False
                            List_File = os.listdir(r"C:\Users\hanbiro\Downloads")
                            for file in List_File :
                                Find_File = (file.rfind(Name))
                                if  Find_File != -1 :
                                    Download = True
                            if  Download == True :
                                commons.WriteOnExcel(data["mail_excel"]["setting"]["folder"]["backup"]["pass"])
                            else :
                                commons.WriteOnExcel(data["mail_excel"]["setting"]["folder"]["backup"]["fail"])
                            break
                                                       
                        i += 1     
                    
                    
                    # Upload Eml   
                    commons.Title("Folders : Upload Eml")           
                    time.sleep(1)
                    commons.ClickElementWithXpath(data["mail"]["folder_upload"])
                    if commons.IsDisplayedByXpath(data["mail"]["folder_file"]) == True :
                        commons.Content("Click on upload emails")
                        
                        Upload_File = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["folder_file"]))) 
                        Upload_File.click()
                        
                        Local_Eml = Param.local + r"\File\automationtest@vndev.hanbiro.com_Mail_1721698949.871362.eml"
                        time.sleep(3)
                        pyautogui.write(Local_Eml)
                        pyautogui.press('enter')
                        driver.find_element_by_css_selector(data["mail"]["folder_file_save"]).click()
                        
                        if commons.IsDisplayedByCss(data["mail"]["folder_file_discard"]) == True :
                            commons.WriteOnExcel(data["mail_excel"]["setting"]["folder"]["upload"]["fail"])
                        else :
                            commons.WriteOnExcel(data["mail_excel"]["setting"]["folder"]["upload"]["pass"])
                            
                        
                    else :
                        commons.CaseFail("Click on upload emails")     
                        commons.WriteOnExcel(data["mail_excel"]["setting"]["folder"]["upload"]["fail"])
                    
                    
                    # Empty Mailbox
                    commons.Title("Folders : Empty Mailbox")
                    commons.ClickElementWithXpath(data["mail"]["folder_empty"])
                    if commons.IsDisplayedByXpath(data["mail"]["folder_empty_save"]) == True :
                        commons.Content("Click on empty mailbox")
                        Empty =  driver.find_element_by_xpath(data["mail"]["folder_empty_save"])
                        Empty.click()
                        
                        if commons.IsDisplayedByXpath(data["mail"]["folder_empty_discard"]) == True :
                            commons.WriteOnExcel(data["mail_excel"]["setting"]["folder"]["empty"]["fail"])
                        else :
                            commons.WriteOnExcel(data["mail_excel"]["setting"]["folder"]["empty"]["pass"])

                    else:
                        commons.CaseFail("Click on empty mailbox")
                        commons.WriteOnExcel(data["mail_excel"]["setting"]["folder"]["empty"]["fail"])
                    
                    
                    # Share
                    commons.Title("Folders : Share")  
                    commons.ClickElementWithXpath(data["mail"]["folder_share"])
                    if  commons.IsDisplayedById(data["mail"]["folder_position"]) == True :
                        commons.Content("Click on buton share")
                        Org = commons.IsDisplayedByClass("ztree")
                        if Org == True :
                            # Total Department #
                            Org_Tree = driver.find_element_by_class_name("ztree")
                            Org_Depart = Org_Tree.find_elements_by_tag_name("li")
                            Org_Depart = commons.TotalData(Org_Depart)
                            i = 1
                            Select = False
                            # Open Department to select user #
                            while i <= Org_Depart :
                                Depart_Close = driver.find_element_by_xpath(data["mail"]["org_open"] % str(i)) 
                                Depart_Close.click()
                                
                                Depart_Sub = driver.find_element_by_xpath(data["mail"]["org_list"] % str(i))
                                Depart_Sub = Depart_Sub.find_elements_by_tag_name("li")
                                Total_Sub  = commons.TotalData(Depart_Sub)
                                j = 1
                                if Total_Sub != 0 :
                                    while j <= Total_Sub :
                                        driver.find_element_by_xpath(data["mail"]["org_list_sub"] % (str(i),str(j))).click()
                                        Is_User = commons.IsDisplayedByXpath(data["mail"]["org_list_user"] % str(j))
                                        if Is_User == True :
                                            Select = True
                                            driver.find_element_by_xpath(data["mail"]["org_user"] % str(j)).click()
                                            commons.Content("Click on checkbox to select user")
                                            driver.find_element_by_css_selector(data["mail"]["org_add"]).click()
                                            commons.Content("Click on icon add user")
                                            driver.find_element_by_css_selector(data["mail"]["org_save"]).click()
                                            commons.Content("Click on button save user")
                                            Height = GetSystemMetrics(1)
                                            time.sleep(2)
                                            driver.execute_script("window.scrollTo(0, %s)" % Height)
                                            driver.find_element_by_xpath(data["mail"]["org_share"])
                                            time.sleep(2)
                                            if commons.IsDisplayedByCss(data["mail"]["org_list_share"]) == True :
                                                commons.WriteOnExcel(data["mail_excel"]["setting"]["folder"]["share"]["pass"])
                                            else :
                                                commons.WriteOnExcel(data["mail_excel"]["setting"]["folder"]["share"]["fail"])
                                            break
                                        j +=1
                                    
                                if  Select == True :
                                    break
                                i +=1
                        
                    else :
                        commons.WriteOnExcel(data["mail_excel"]["setting"]["folder"]["share"]["fail"])
                    
                    
                    
                    # Modify Folder #          
                    commons.Title("Folders : Modify")  
                    Modify = False
                    Folder_Modify = Folder + "[Modified]"
                    Input_Modify = driver.find_element_by_css_selector(data["mail"]["folder_input"])
                    Input_Modify.clear()
                    Input_Modify.send_keys(Folder_Modify)
                    commons.Content("Input folder name modify")
                    driver.find_element_by_xpath(data["mail"]["folder_save"]).click()
                    commons.Content("Click on button save folder modify")
                    
                    No_Data = commons.IsDisplayedByXpath(data["mail"]["folder_no"])
                    if  No_Data == True :
                        commons.WriteOnExcel(data["mail_excel"]["setting"]["folder"]["modify"]["fail"])
                    else :
                        Check = False
                        i = 1
                        while i <= List_Folder :
                            Name_Driver = driver.find_element_by_xpath(data["mail"]["folder_name"] % str(i))
                            Name_Modify = Name_Driver.text
                            Name_Modify = Name_Modify[int(Name_Modify.rfind("F")): int(Name_Modify.rfind("("))]
                            if  Name_Modify == Folder_Modify :
                                Check = True
                                break
                            i += 1
                        if  Check == True :
                            commons.WriteOnExcel(data["mail_excel"]["setting"]["folder"]["modify"]["pass"])
                        else :
                            commons.WriteOnExcel(data["mail_excel"]["setting"]["folder"]["modify"]["fail"])
                    
                    
                    # Delete Folder #         
                    commons.Title("Folders : Delete") 
                    Delete = False
                    i = 1
                    while i <= List_Folder :
                        Folder_Delete = driver.find_element_by_xpath(data["mail"]["folder_name"] % str(i))
                        Name_Delete = Folder_Delete.text
                        Name_Delete = Name_Delete[int(Name_Modify.rfind("F")): int(Name_Modify.rfind("]"))+1]
                        
                        if  Name_Delete == Folder or  Name_Delete == Folder + "[Modified]":
                            Delete = True
                            Name_Driver.click()
                            commons.Content("Click on folder name to delete")
                            break
                        i += 1
                    if  Delete == True :
                        driver.find_element_by_xpath(data["mail"]["folder_delete"]).click()
                        commons.Content("Click on button delete folder")
                        if List_Folder == 1 :
                            No_Data = commons.IsDisplayedByXpath(data["mail"]["folder_no"])
                            if No_Data == True :
                                commons.WriteOnExcel(data["mail_excel"]["setting"]["folder"]["delete"]["pass"])
                        else :
                            i = 1
                            Check = True
                            time.sleep(3)
                            List_Folder =  driver.find_elements_by_xpath(data["mail"]["folder_item"])
                            List_Folder =  commons.TotalData(List_Folder)
                            List_Folder = List_Folder / 2
                            while i <= List_Folder :
                                Folder_Delete = driver.find_element_by_xpath(data["mail"]["folder_name"] % str(i))
                                Name_Delete = Folder_Delete.text
                                Name_Delete = Name_Delete[int(Name_Modify.rfind("F")): int(Name_Modify.rfind("]"))+1]
                                
                                if  Name_Delete == Folder or  Name_Delete == Folder + "[Modified]":
                                    Check = False
                                    break
                                i += 1
                            if  Check == True :
                                commons.WriteOnExcel(data["mail_excel"]["setting"]["folder"]["delete"]["pass"])
                            else :
                                commons.WriteOnExcel(data["mail_excel"]["setting"]["folder"]["delete"]["fail"])
                            
                    
                if  Create == False :
                    commons.WriteOnExcel(data["mail_excel"]["setting"]["folder"]["create"]["fail"])
                
            else :
                commons.WriteOnExcel(data["mail_excel"]["setting"]["folder"]["create"]["fail"])

def Forwarding():
    commons.Title("Forwarding")
    commons.ClickElementWithXpath(data["mail"]["forwar"])
    if  commons.IsDisplayedByCss(data["mail"]["forwar_add"]) == True :
        Access = True 
        commons.WriteOnExcel(data["mail_excel"]["setting"]["forwar"]["access"]["pass"])
    else :
        Access = False
        commons.WriteOnExcel(data["mail_excel"]["setting"]["forwar"]["access"]["fail"])
    
    if  Access == True :
        # Create #
        commons.Title("Forwarding : Add")
        driver.find_element_by_css_selector(data["mail"]["forwar_add"]).click()
        if  commons.IsDisplayedByXpath(data["mail"]["forwar_save"]) == True :
            Add = True
            commons.Content("Click on button add")
        else :
            commons.CaseFail("Click on button add")
            Add = False
            commons.WriteOnExcel(data["mail_excel"]["setting"]["forwar"]["create"]["fail"])
        if  Add == True :
            Create = False
            driver.find_element_by_id(data["mail"]["forwar_task"]).send_keys(data["mail"]["forwar_text"])
            commons.Content("Input email")
            driver.find_element_by_xpath(data["mail"]["forwar_save"]).click()
            commons.Content("Click on button save")
            
            
            if commons.IsDisplayedByXpath(data["mail"]["forwar_no"]) == True :
                commons.WriteOnExcel(data["mail_excel"]["setting"]["forwar"]["create"]["fail"])
            else :
                Table = driver.find_element_by_xpath(data["mail"]["forwar_table"])
                Total =  Table.find_elements_by_tag_name("tr")
                Total =  commons.TotalData(Total)
                Count = 0
                i = 1
                while i <= Total :
                    Email = driver.find_element_by_xpath(data["mail"]["forwar_name"] % str(i)).text
                    if Email == data["mail"]["forwar_email_1"] or Email == data["mail"]["forwar_email_2"]:
                        Count += 1
                    i +=1
                if  Count == 2 :
                    Create = True
                    commons.WriteOnExcel(data["mail_excel"]["setting"]["forwar"]["create"]["pass"])
                else :
                    commons.WriteOnExcel(data["mail_excel"]["setting"]["forwar"]["create"]["fail"])
                    
            # Delete # 
            commons.Title("Forwarding : Delete") 
            if commons.IsDisplayedByXpath(data["mail"]["forwar_no"]) == True :
                commons.Content("No Forwarding to delete")
            else :
                if Create == True :
                    i = Total
                    Count = 0
                    while i >= 1 :
                        Email = driver.find_element_by_xpath(data["mail"]["forwar_name"] % str(i)).text
                        if  Email == data["mail"]["forwar_email_1"] or Email == data["mail"]["forwar_email_2"]:
                            driver.find_element_by_xpath(data["mail"]["forwar_checkbox"] % str(i)).click()
                            commons.Content("Click on checkbox")
                            driver.find_element_by_xpath(data["mail"]["forwar_delete"]).click()
                            commons.Content("Click on button delete")
                            Count += 1
                        i -=1

                    if Count != 0 :
                        if commons.IsDisplayedByXpath(data["mail"]["forwar_no"]) == True :
                            if  Total == Count :
                                commons.WriteOnExcel(data["mail_excel"]["setting"]["forwar"]["delete"]["pass"])
                            else :
                                commons.WriteOnExcel(data["mail_excel"]["setting"]["forwar"]["delete"]["fail"])
                        else :
                            Table = driver.find_element_by_xpath(data["mail"]["forwar_table"])
                            Total_After =  Table.find_elements_by_tag_name("tr")
                            Total_After =  commons.TotalData(Total_After)
                            if Total == Total_After + Count :
                                commons.WriteOnExcel(data["mail_excel"]["setting"]["forwar"]["delete"]["pass"])
                            else :
                                commons.WriteOnExcel(data["mail_excel"]["setting"]["forwar"]["delete"]["fail"])


def Whitelist():
    commons.Title("Whitelist")
    commons.ClickElementWithXpath(data["mail"]["white"])

    if  commons.IsDisplayedById(data["mail"]["white_input"]) == True :
        Access = True 
        commons.WriteOnExcel(data["mail_excel"]["setting"]["white"]["access"]["pass"])
    else :
        Access = False
        commons.WriteOnExcel(data["mail_excel"]["setting"]["white"]["access"]["fail"])
    
    if  Access == True :
        
        if commons.IsDisplayedByXpath(data["mail"]["white_no"]) == False :
            
            Table = driver.find_element_by_xpath(data["mail"]["white_table"])
            Total =  Table.find_elements_by_tag_name("tr")
            Total =  commons.TotalData(Total)
            
            i = 1
            while i <= Total :
                Email = driver.find_element_by_xpath(data["mail"]["white_name"] % str(i)).text
                if  Email == data["mail"]["white_email"] :
                    commons.Title("Whitelist : Delete")
                    driver.find_element_by_xpath(data["mail"]["white_checkbox"] % str(i)).click()
                    commons.Content("Click on checkbox")
                    driver.find_element_by_xpath(data["mail"]["white_delete"]).click()
                    commons.Content("Click on button delete")

                    if  Total == 1 :
                        if commons.IsDisplayedByXpath(data["mail"]["white_no"]) == True :
                            commons.WriteOnExcel(data["mail_excel"]["setting"]["white"]["delete"]["pass"])
                        else :
                            commons.WriteOnExcel(data["mail_excel"]["setting"]["white"]["delete"]["fail"])
                    else :
                        Delete = False
                        j = 1
                        try :
                            while j <= Total :
                                Email = driver.find_element_by_xpath(data["mail"]["white_name"] % str(j)).text
                                if Email == data["mail"]["white_email"] :
                                    Delete = True
                                    commons.WriteOnExcel(data["mail_excel"]["setting"]["white"]["delete"]["fail"])
                                    break
                                j += 1
                            if  Delete == True :
                                commons.WriteOnExcel(data["mail_excel"]["setting"]["white"]["delete"]["fail"])
                        except :
                            pass
                i +=1
        
        
        commons.Title("Whitelist : Add")     
        driver.find_element_by_id(data["mail"]["white_input"]).send_keys(data["mail"]["white_email"])
        commons.Content("Input Whitelist")
        
        driver.find_element_by_css_selector(data["mail"]["white_add"]).click()
        commons.Content("Click on button add")
        
        if commons.IsDisplayedByXpath(data["mail"]["white_no"]) == True :
            commons.WriteOnExcel(data["mail_excel"]["setting"]["white"]["create"]["fail"])
        else :
            Table = driver.find_element_by_xpath(data["mail"]["white_table"])
            Total =  Table.find_elements_by_tag_name("tr")
            Total =  commons.TotalData(Total)
            
            i = 1
            Create = False
            while i <= Total :
                Email = driver.find_element_by_xpath(data["mail"]["white_name"] % str(i)).text
                if Email == data["mail"]["white_email"] :
                    Create = True
                    commons.WriteOnExcel(data["mail_excel"]["setting"]["white"]["create"]["pass"])
                    break
                i +=1
            
            if  Create == False :
                commons.WriteOnExcel(data["mail_excel"]["setting"]["white"]["create"]["fail"])


def SmtpPop3Imap():
    commons.Title("SMTP-POP3-IMAP")
    commons.ClickElementWithXpath(data["mail"]["smtp"])
    if  commons.IsDisplayedByXpath(data["mail"]["smtp_save"]) == True :
        Access = True 
        commons.WriteOnExcel(data["mail_excel"]["setting"]["smtp"]["access"]["pass"])
    else :
        Access = False
        commons.WriteOnExcel(data["mail_excel"]["setting"]["smtp"]["access"]["fail"])

def AliasAccounts():
    commons.Title("Alias Accounts")
    commons.ClickElementWithXpath(data["mail"]["alias"])
    if  commons.IsDisplayedById(data["mail"]["alias_search"]) == True :
        Access = True 
        commons.WriteOnExcel(data["mail_excel"]["setting"]["alias"]["access"]["pass"])
    else :
        Access = False
        commons.WriteOnExcel(data["mail_excel"]["setting"]["alias"]["access"]["fail"])


def AutoCompleteSetting():
    commons.Title("Auto-Complete Setting")
    commons.ClickElementWithXpath(data["mail"]["comp"])
    if  commons.IsDisplayedByXpath(data["mail"]["comp_reset"]) == True :
        Access = True 
        commons.WriteOnExcel(data["mail_excel"]["setting"]["comp"]["access"]["pass"])
    else :
        Access = False
        commons.WriteOnExcel(data["mail_excel"]["setting"]["comp"]["access"]["fail"])

def MailFetching():
    commons.Title("Mail Fetching")
    commons.ClickElementWithXpath(data["mail"]["fetching"])
    if  commons.IsDisplayedByXpath(data["mail"]["fetching_pass"]) == True :
        Access = True 
        commons.WriteOnExcel(data["mail_excel"]["setting"]["comp"]["access"]["pass"])
    else :
        Access = False
        commons.WriteOnExcel(data["mail_excel"]["setting"]["comp"]["access"]["fail"])

def Settings():
    commons.Title("Settings")
    # Frame 2
    driver.switch_to.frame(commons.FindElementById("newMailIframe"))
    
    driver.find_element_by_xpath(data["mail"]["setting"]).click()
    if commons.IsDisplayedByXpath(data["mail"]["signature"]) == True :
        Settings = True
        commons.WriteOnExcel(data["mail_excel"]["setting"]["access"]["pass"])
    else :
        Settings = False
        commons.WriteOnExcel(data["mail_excel"]["setting"]["access"]["fail"])
    
   
    
    
    if  Settings == True :
        
        try :
            WritingSetting()
        except :
            commons.CaseFail("Error Xpath : Writing Setting")
        
        try :            
            Signature()
        except:
            commons.CaseFail("Error Xpath : Signature")
            
        try :            
            AutoSort()
        except:
            commons.CaseFail("Error Xpath : Auto Sort")
       
        try :
            VacationAutoReplies()
        except:
            commons.CaseFail("Error Xpath : Vacation Auto Replies")
        
        try :
            BlockedAddresses()
        except:
            commons.CaseFail("Error Xpath : Blocked Addresses")
            
        try :
            Folders()
        except:
            commons.CaseFail("Error Xpath : Folders")
        
        try :
            Forwarding()
        except:
            commons.CaseFail("Error Xpath : Forwarding")
        
        try :
            Whitelist()
        except:
            commons.CaseFail("Error Xpath : Whitelist")
            
        try :
            SmtpPop3Imap()
        except:
            commons.CaseFail("Error Xpath : SMTP POP3 IMAP")
            
        try :
            AliasAccounts()
        except:
            commons.CaseFail("Error Xpath : Alias Accounts")
            
        try :
            AutoCompleteSetting()
        except:
            commons.CaseFail("Error Xpath : Auto Complete Setting")
        
        try :
            MailFetching()
        except:
            commons.CaseFail("Error Xpath : Mail Fetching")
        
    # Close Frame 3 - Frame open from def Signature()
    commons.SwitchToDefaultContent()
    
    
def MenuMail():
    
    Access = AccessMail() 
    if  Access == True :
        Settings()
    
    