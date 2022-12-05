try:
    from datetime import datetime as dt
    start_time = dt.now()
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders
    import csv
    from numpy import NaN
    import pandas as pd
    import datetime 
    from random import randint
    from time import sleep
    from platform import python_version
    try:
        def attendance_report():
            #reading the input files using pandas.
            dataframe_attendance = pd.read_csv('input_attendance.csv')
            dataframe_registered = pd.read_csv('input_registered_students.csv')
            dataframe_attendance['Roll']=''
            dataframe_attendance['Name']=''
            deleted_rows = 0
            #Here i have made two separate columns for roll no,name which is divided from attednace column,later dropping the attendance column.
            for i in range(0,len(dataframe_attendance)):
                x = dataframe_attendance['Attendance'][i]
                if(type(x)==str):
                    y = x[0:8]
                    z = x[9:]
                    dataframe_attendance['Roll'][i]=y
                    dataframe_attendance['Name'][i]=z
                else:
                    deleted_rows+=1
                    dataframe_attendance['Roll'][i]=NaN
                    dataframe_attendance['Name'][i]=NaN
            dataframe_attendance.drop('Attendance',inplace=True,axis=1)
            #here aslo in the same way from the timestamp i am seprating both date and time using pandas.
            dataframe_attendance['Date']=''
            dataframe_attendance['Time']=''
            for i in range(0,len(dataframe_attendance)):
                x = pd.to_datetime(dataframe_attendance['Timestamp'][i],format = "%d-%m-%Y %H:%M")
                dataframe_attendance['Date'][i]=x.date()
                dataframe_attendance['Time'][i]=x.time()
                dataframe_attendance['Timestamp'][i]=x

            for j in range(0,len(dataframe_attendance)):
                start_date = dataframe_attendance['Timestamp'][j]
                break
            for j in range(0,len(dataframe_attendance)):
                end_date = dataframe_attendance['Timestamp'][j]  
            #lecture days are the days between start_date and end_date which includes all mondays and thursdays.
            #created this dictonaries to store no of lecture days between start date and end date.
            #lecture_days contains only valid dates(means mon,thur)
            lecture_days =  {}
            #total_days contains all the dates between start_date to end_date..but (mon,thur marked as 1, other marked as 0)
            total_days = {}
            while(start_date.date()<=end_date.date()):
                if(start_date.day_name()=="Monday" or start_date.day_name()=='Thursday'):
                    lecture_days[start_date.date()]=1
                    total_days[start_date.date()]=1
                else:
                    total_days[start_date.date()]=0
                #using pandas dateoffset i am iterating each date.
                start_date = start_date + pd.DateOffset(days=1)
            #now i am sorting the dataframe_attendance based on rol no date time so it would be convenient to compare with data of registered_students.
            sorted_df = dataframe_attendance.sort_values(by=['Roll','Date','Time'])
            sorted_df = sorted_df.reset_index()
            sorted_df.drop('index',inplace=True,axis=1)
            sorted_df.drop('Timestamp',inplace=True,axis=1)
            #storing the lecture dates in the list called dates_list.
            dates_list = []

            #mapping the dates dict help full to give unique id for each date.
            # like date 1,date 2..date n.
            mapping_dates = {}
            count = 1
            for x in lecture_days:
                dates_list.append(x)
                mapping_dates[x]=count
                count+=1

            #converting dataframe_registered dataframe into such a way that its output should match attendance_report_consolidated.
            for x in range(0,len(dates_list)):
                dataframe_registered['Date '+str(x+1)]='A'
            dataframe_registered['Actual Lecture Taken']=''
            dataframe_registered['Total Real']=''
            dataframe_registered['%Attendance']=''
            #actual lecture taken....means total no of lectures taken by sir.
            #total real ...the no of lectures where student is present.
            pres_row = 0

            for i in range(0,len(dataframe_registered)):
                x = dataframe_registered['Roll No'][i]
                #storing the present roll number in the x.
                #temp_df is creating temporary columns using dataframe(pandas).
                temp_df = pd.DataFrame(columns = ['Date','Roll','Name','Total Attendance Count','Real','Duplicate','Invalid','Absent'])
                #creating initial dict for each date.
                some_dicts = {'Date':'','Roll':x,'Name':dataframe_registered['Name'][i],'Total Attendance Count':0,'Real':0,'Duplicate':0,'Invalid':0,'Absent':0}
                #making dataframe of above dict.
                row_dataframe = pd.DataFrame(some_dicts,index = [0])
                #concating temp_df with some_dicts.
                temp_df = pd.concat([temp_df,row_dataframe],ignore_index = True)
                #dict_fake stores the date as 1, if it is not valid date.
                dict_fake = {}
                #dict_pres stores the date as 1, if it is valid date.
                dict_pres = {}
                for p in range(0,len(dates_list)):
                    dict_fake[dates_list[p]]=0
                    dict_pres[dates_list[p]]=0
                #now iterating for each roll no through attendance_sheet dataframe till it matches with its date.
                for j in range(pres_row,len(sorted_df)):
                    y = sorted_df['Roll'][j]
                    if(type(y)==float):
                        break
                    if(x<y):
                        pres_row=j
                        break
                    else :
                        #total_days[date]==1 is one means it is either mon or thursday..
                        #and timeshould be between 2 to 3 to be a valid date.
                        if(total_days[sorted_df['Date'][j]]==1 and  str(sorted_df['Time'][j])>='14:00:00' and str(sorted_df['Time'][j])<='15:00:00'):
                            dataframe_registered["Date "+ str(mapping_dates[sorted_df['Date'][j]])][i]='P'
                            #so keep it as P, also make dict_pres = 1 as he attended the class.
                            dict_pres[sorted_df['Date'][j]]+=1
                        elif(total_days[sorted_df['Date'][j]]==1):
                            #it is mon or thur but not in between 2 to 3, so it comes under fake.
                            dict_fake[sorted_df['Date'][j]]+=1
                no_of_ps = 0
                #here in this variable we store no of present in the lecture dates
                for z in range(0,len(dates_list)):
                    if dataframe_registered['Date '+str(z+1)][i]=='P':
                        no_of_ps+=1
                dataframe_registered['Actual Lecture Taken'][i]=len(dates_list)
                dataframe_registered['Total Real'][i]=no_of_ps
                dataframe_registered['%Attendance'][i]=round((no_of_ps/len(dates_list))*100,2)
                # s = pd.Series((len(dataframe_registered)-13)*[None,None,None],index=['Date','Roll','Name','Total Attendance Count','Real','Duplicate','Invalid','Absent'])
                #dict_fake contains the invalid dates or fake attedances.
                for count1 in range(1,len(dates_list)+1):
                    temp_dict_perdate = {}
                    #this will store the invalids or absent or real for particular date of each person
                    temp_dict_perdate['Date'] = "Date " + str(count1)
                    if (dict_pres[dates_list[count1-1]]>0):
                        temp_dict_perdate['Real'] = 1
                    else:
                        temp_dict_perdate['Real'] = 0
                    if (dict_pres[dates_list[count1-1]]>0):
                        temp_dict_perdate['Duplicate'] = dict_pres[dates_list[count1-1]]-1
                    else:
                        temp_dict_perdate['Duplicate'] = 0
                    temp_dict_perdate['Invalid'] = dict_fake[dates_list[count1-1]]
                    if (dict_pres[dates_list[count1-1]]>0):
                        temp_dict_perdate['Absent'] = 0
                    else:
                        temp_dict_perdate['Absent'] = 1
                    #as we know total attedance count = sum of real + duplicate + invalid
                    temp_dict_perdate['Total Attendance Count'] = temp_dict_perdate['Real']+temp_dict_perdate['Duplicate']+temp_dict_perdate['Invalid']

                    df2 = pd.DataFrame(temp_dict_perdate,index=[0])
                    #concating the df2..temp_dict_perdate each time to the main student data.
                    temp_df = pd.concat([temp_df,df2],ignore_index=True)
                #now here i am counting all total attendance count,real,duplicate,invalid,absent of a each student by adding the data of individual
                #dates.
                for count2 in range(1,len(dates_list)+1):
                    temp_df.loc[0,'Total Attendance Count'] += temp_df.loc[count2,'Total Attendance Count']
                    temp_df.loc[0,'Real'] += temp_df.loc[count2,'Real']
                    temp_df.loc[0,'Duplicate'] += temp_df.loc[count2,'Duplicate']
                    temp_df.loc[0,'Invalid'] += temp_df.loc[count2,'Invalid']
                    temp_df.loc[0,'Absent'] += temp_df.loc[count2,'Absent']
                #now converting this dataframe into excel sheet with output file according to roll number.
                temp_df.to_excel('./output/'+  str(x) + '.xlsx',index=False)
            #complete dataframe_registered will be the dataframe of attendance_report_consolidated file.
            dataframe_registered.to_excel('./output/attendance_report_consolidated.xlsx',index=False)
    except:
        print('Any compilation error would be in the function')
    
    # Python code to illustrate Sending mail with attachment.
    try:
        def send_mail():
            #here add your gmail email-adress.
            fromaddr = "srivardhanrao9009@gmail.com"
            #add the sender's gmail adress.
            toaddr = "srins4545@gmail.com"

            # instance of MIMEMultipart
            msg = MIMEMultipart()

            #here i am storing the sender's email and receivers email address
            msg['From'] = fromaddr
            msg['To'] = toaddr
           
            # storing the subject 
            msg['Subject'] = "Tut 06 attendance_report_consolidated file"

            #here i am storing the body of the mail,in the string.
            body = "here is the file"

            # attach the body with the msg instance
            msg.attach(MIMEText(body, 'plain'))

            # add the file name with the extension here which you want to sent.
            filename = "attendance_report_consolidated.xlsx"
            attachment = open("./output/attendance_report_consolidated.xlsx", "rb")

            # instance of MIMEBase and named as p
            p = MIMEBase('application', 'octet-stream')

            # To change the payload into encoded form
            p.set_payload((attachment).read())

            # encode into base64
            encoders.encode_base64(p)

            p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
            msg.attach(p)

            s = smtplib.SMTP('smtp.gmail.com', 587)

            # for security purpose start the tls.
            try:
                s.starttls()
            except:
                print('Please dont use IIT Patna wifi, use personal hotspot to send the mail.')
                return
            #here use your email adress, and authenication password.(enable 2-step verification password, required to send the mail)
            s.login(fromaddr, "gmuuknfpuxnlfkif")

            #now convert multipart msg into the string.
            text = msg.as_string()

            # here i am sending the email..by representing fromaddr, toaddr, text.
            s.sendmail(fromaddr, toaddr, text)

            #quitting the session atlast.
            s.quit()
    except:
        print('some compilation error inside the send mail func')

    attendance_report()
    print('output files generated')
    send_mail()
    ver = python_version()

    if ver == "3.8.10":
        print("Correct Version Installed")
    else:
        print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")
    #This shall be the last lines of the code.
    end_time = dt.now()
    print('Duration of Program Execution: {}'.format(end_time - start_time))
except:
    print("Please install the required libraries.")












