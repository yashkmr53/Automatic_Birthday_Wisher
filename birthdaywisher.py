import pandas as pd
import datetime as dt
import smtplib
import os

#change the directory to source code folder to use windows automatic task scheduler
os.chdir(r"C:\Users\yashk\PythonVScode")

mail_id='Enter your email'
password="Enter your password"

def sendEmail(to,sub,msg):
    s=smtplib.SMTP(host="smtp.gmail.com",port=587)
    s.starttls()
    s.login(user=mail_id,password=password)
    s.sendmail(from_addr=mail_id, to_addrs=to, msg=f"Subject : {sub}\n\n {msg}")
    s.quit()
    print('sucess')
    pass

if __name__ == "__main__":

    # read excel data by pandas
    df=pd.read_excel("birthdaydata.xlsx")

    # read today's date by datetime with the help of strftime to check birth date (dd-mm)
    today=dt.datetime.now().strftime("%d-%m")

    # read current year to compare in dataset so that extra wishes does not happen
    yearNow=dt.datetime.now().strftime("%Y")

    # iterating dataset with itrrows() so that i can able to access each values as row_index and values as series(dictonary type)
    for index, value in df.iterrows():

        # compare birthdate and year of the dataset with (d-m) format
        if (value['Birthday'].strftime('%d-%m'))==today and yearNow not in str(value['Year']) :
           
            # call send email function by passing mailId,subject and msg
            to=value['Email']
            sub="Happy Birthday"
            msg=value['Comment']
            sendEmail(to,sub,msg)

            # update Year of the dataset with current year so that it can not wish again using loc fuction of df
            df.loc[index,'Year']=str(value['Year'])+','+str(yearNow)

            # overrite the same file with index=False because we do not want new columns in the dataset
            df.to_excel("birthdaydata.xlsx",index=False)
