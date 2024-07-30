from gw_v3_login import Login 
from gw_v3_mail import MenuMail 
from gw_v3_board import MenuBoard

def run(domain,user,password) :
    Check_login = Login(domain,user,password)
    if  Check_login == True :
        #MenuMail()
        MenuBoard()
        
        
        
run("vndev.hanbiro.com","automationtest","automationtest1!")

















