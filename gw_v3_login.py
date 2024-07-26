from gw_v3_set_up import driver,data,Keys,EC,By,WebDriverWait,json
from gw_v3_functions import commons, time

def Login(domain,user,password):
    commons.Title("LOGIN")
    driver.get(data["domain"] % domain)
    
    # Input Id
    Id = commons.IsDisplayedByIdLogin()
    if  Id != False :
        commons.FindElementById(":r%s:" % Id).send_keys(user)
        commons.Content("Input ID")
        
    # Input Pass 
    commons.FindElementById("gw_pass").send_keys(password)
    commons.Content("Input Pass")
    
    # Click Btn
    if  Id != False :
        commons.FindElementById(":r%s:" % Id).send_keys(Keys.RETURN)   
        commons.Content("Click on button login")
        
    return True
    
    


 