from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from time import sleep
import urllib.parse
from tkinter import *
from tkinter import scrolledtext,messagebox
import tkinter.font
from tkinter import filedialog
import pandas as pd
from PIL import *
import PIL

driver=None
#driver = webdriver.Chrome(executable_path=r'C:\Users\Dell\Desktop\Whatsapp\chromedriver_win32\chromedriver.exe')
Link = "https://web.whatsapp.com/"
wait = None




def whatsapp_login(nums,msg,stbrow,stmsg):
    global wait, driver, Link
    if(nums != '' and msg != ''):
        chrome_options = Options()
        chrome_options.add_argument('--user-data-dir=./User_Data')
        driver = webdriver.Chrome(options=chrome_options)
        wait = WebDriverWait(driver, 20)
        print("SCAN YOUR QR CODE FOR WHATSAPP WEB IF DISPLAYED")
        driver.get(Link)
        driver.maximize_window()
        print("QR CODE SCANNED")
        print(msg)
        numlst=nums.split('\n')
        
        #print(numlst)
        if(numlst[-1]==''):
            numlst.pop()
        for x in numlst:
            send_message('91'+x,msg, 1)
            sleep(5)
        driver.close() # Close the Open tab
        driver.quit()
        stbrow.delete(1.0, END)
        stmsg.delete(1.0, END)
        messagebox.showinfo("Whatsapp Notification", "Messages Sent Sucessfully")
    else:
        messagebox.showinfo("Whatsapp Notification", "Empty Field!")
    
    
    
def send_message(number,msg,count):
    print("In send_message_to_unsavaed_contact method")
    params = {'phone': str(number), 'text': str(msg)}
    end = urllib.parse.urlencode(params)
    final_url = Link + 'send?' + end
    print(final_url)
    driver.get(final_url)
    WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.XPATH, '//div[@title = "Menu"]')))
    print("Page loaded successfully.")
    for retry in range(3):
        try:
            sleep(5)
            wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/footer/div[1]/div[3]/button'))).click()
            break
        except Exception as e:
            print(number+' is not a valid Whatsapp Number')

            break
            #print("Fail during click on send button.")
            '''if retry==2:return
    msg_box = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')
    for index in range(count-1):
        msg_box.send_keys(msg)
        driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[3]/button').click()'''
    print("Message sent successfully.")
    
def browse(stbrow):
    global data
    
    import_file_path = filedialog.askopenfilename()
    data = pd.read_excel (import_file_path)
    phlist = data["Phone No"].tolist()
    for i in phlist:
        stbrow.insert(END,i)
        stbrow.insert(END,'\n')




#GUI    
root=Tk()
root.geometry('900x720+350+50')
root.title("Whatsapp Messenger made by Madhavi Bisen")
root.resizable(0,0)
root.config(bg='spring green')

#Heading
lblhead=Label(root,text='Whatsapp Messenger Using Python',font=('arial',23,'bold'),fg='black',bg='spring green',width=30)
lblhead.place(x=170,y=10)

#mainframe
centerFrame=Frame(root,bg='white')
centerFrame.place(x=100,y=95,width=700,height=600)

FontOfEntryList=tkinter.font.Font(family="TimesNewRoman",size=15)
FontOfDropList=tkinter.font.Font(family="TimesNewRoman",size=10)

btnbrow=Button(centerFrame,bg='white',text="Select Contact",bd=5,width=13,font=("Aerial",12,"bold"),command=lambda: browse(stbrow))
btnbrow.place(x=150,y=100)
    
stbrow=scrolledtext.ScrolledText(centerFrame,bd=3,font=FontOfEntryList,width='35',height='3')
stbrow.place(x=150,y=150)

lblmsg=Label(centerFrame,text="Type Your Message : ",font=('Aerial',20,'bold'),bg='white',fg='black')
lblmsg.place(x=150, y=300)

stmsg=scrolledtext.ScrolledText(centerFrame,bd=3,font=FontOfEntryList,width='35',height='6')
stmsg.place(x=150,y=350)

btnsend=Button(centerFrame,bg='white',text="Send",bd=5,width=10,font=("Aerial",12,"bold"),command=lambda: whatsapp_login(stbrow.get('1.0', 'end-1c'),stmsg.get('1.0', 'end-1c'),stbrow,stmsg))
btnsend.place(x=290,y=520)

                                                                                                
root.mainloop()