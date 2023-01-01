import win32com.client
import pyodbc
import pandas as pd
from datetime import datetime
import glob
import time



time.sleep(5)


conn=pyodbc.connect('Driver={SQL Server};'
                      'Server=XXXX;'
                      'Database=XXXX;'
                      'Integrated Security=False;'
                      'uid=XXXX;'
                        'pwd=XXXX;'
                        
                              )

PathtoFolder=r'C:\Users\odemirhan\TURISTIK HAVA TASIMACILIK A.S\Gökmen Düzgören - FOE_2019\phyton\db_python\FDE\ACARSmessages'
outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application").GetNamespace("MAPI")
          
FDEfolder= outlook.GetDefaultFolder(6).Folders.Item("MAX FDE Messages")

#ArchieveFDE= outlook.GetDefaultFolder(6).Folders.Item("FDE Archieves")
messages = FDEfolder.Items
MSGdf=pd.read_excel(r"C:\Users\odemirhan\TURISTIK HAVA TASIMACILIK A.S\Gökmen Düzgören - FOE_2019\phyton\db_python\FDE/MSG.xlsx")
FDEdf=pd.read_excel(r"C:\Users\odemirhan\TURISTIK HAVA TASIMACILIK A.S\Gökmen Düzgören - FOE_2019\phyton\db_python\FDE/FIM.xlsx")



def isdate(datee, timee):
    try:
        timee=timee.replace("AIRPORT","")
        timee=timee.replace("POLICY","")
        dt=datetime.strptime(datee+" "+timee, "%d%b%y %H:%M:%S")
        return True
    except:
        return False


#print(len(messages))
for cnt0 in range(len(messages)):
    
    AIRCRAFT=""
    CALLSIGN=""
    AIRPORTS=""
    TOHOUR=""
    TODATE=""
    TODT=""

    Indexnumber=glob.glob(PathtoFolder+"/*.msg")
    if not len(Indexnumber)==0:
        Indexnumber=[x.replace('C:\\Users\\odemirhan\\TURISTIK HAVA TASIMACILIK A.S\\Gökmen Düzgören - FOE_2019\\phyton\\db_python\\FDE\\ACARSmessages\\','') for x in Indexnumber]
        Indexnumber=[x.replace('.msg','') for x in Indexnumber]
        Indexnumber=[int(x) for x in Indexnumber]
        MailIndex=max(Indexnumber)+1
    else:
        MailIndex=10001
    
    
    
   
    
    wholebody=messages[cnt0+1].Body
    if 'FDE ' in wholebody:
        wholebody=wholebody.replace('\n',' ')
        wholebody=wholebody.replace('\r',' ')
        splittedbody=wholebody.split(" ")
        splittedbody=[x for x in splittedbody if x != '']
        splittedbody =[x.replace('\n','') for x in splittedbody]
        splittedbody =[x.replace('\r','') for x in splittedbody]
        #print(splittedbody)
        #if 'A' in splittedbody:
            #splittedbody.remove('A')
        #if '/' in splittedbody:
            #splittedbody.remove('/')
        for cnt1 in range(len(splittedbody)):
            
            if splittedbody[cnt1]=="PLF" or splittedbody[cnt1]=="RTE":
                AIRCRAFT=splittedbody[cnt1+4]
                CALLSIGN=splittedbody[cnt1+5]
                AIRPORTS=splittedbody[cnt1+6]
                TOHOUR=splittedbody[cnt1+10]
                TODATE=splittedbody[cnt1+11]
                if CALLSIGN=="/":
                    CALLSIGN=""
                    AIRPORTS=""
                    
                try:
                    TODT=datetime.strptime(TODATE+" "+TOHOUR, "%d%b%y %H%M")
                except:
                    TODT=datetime.strptime(splittedbody[cnt1+10]+" "+splittedbody[cnt1+9], "%d%b%y %H%M")
                
            elif splittedbody[cnt1].strip()=="FDE" or splittedbody[cnt1].strip()=="MSG":
                
                FDEorMSG=""
                CODE=""
                OCCURANCEHR=""
                OCCURANCEDATE=""
                OCCURANCEDT=""
                
                                                    
                FDEorMSG=splittedbody[cnt1].strip()
                CODE=splittedbody[cnt1+1].strip()
                try:
                    if FDEorMSG=="FDE":
                        FDEdummy=FDEdf[FDEdf["Fault_Code"].astype(str)==CODE]
                        FAULT=FDEdummy.iat[0,1]
                    else:
                        MSGdummy=MSGdf[MSGdf["Number"].astype(str)==CODE]
                        FAULT=MSGdummy.iat[0,1]
                except:
                    FAULT=""
                    
                
                try:
                    OCCURANCEHR=splittedbody[cnt1+3]
                except:
                    OCCURANCEHR=""
                try:
                    OCCURANCEDATE=splittedbody[cnt1+4]
                except:
                    OCCURANCEDATE=""
                    
                try:
                    OCCURANCEDT=datetime.strptime(OCCURANCEDATE+" "+OCCURANCEHR, "%d%b%y %H%M")
                except:
                    try:
                        OCCURANCEDT=datetime.strptime(splittedbody[cnt1+3]+" "+splittedbody[cnt1+2], "%d%b%y %H%M")
                    except:
                        OCCURANCEDT=""
                
                
                cur1=conn.cursor()
                cur12=conn.cursor()
                listtowrite=(TODT,OCCURANCEDT, AIRCRAFT, CALLSIGN, AIRPORTS, FDEorMSG, CODE )
                cur1.execute("SELECT * from dbo.[FDE_MSG] WHERE [Mail Index]=? AND Aircraft=? AND Code=? AND [Occurance Datetime]=? ", (MailIndex, AIRCRAFT,CODE,OCCURANCEDT))
                ftcur=cur1.fetchone()
                print(ftcur)
                if not ftcur:
                    
                    results=cur12.execute("SELECT TOP 1 * from dbo.[FDE_MSG] WHERE [Mail Index]=? ORDER BY [Occurance Index] Desc", (MailIndex))
                    ftcur2=results.fetchone()
                    #print(ftcur2)
                    if not ftcur2:
                        indexer=1
                    else:
                        row_to_list = [elem for elem in ftcur2]
                        #print(row_to_list)
                        indexer=int(row_to_list[1])+1
                    
                    
                
                    listtowrite2=(MailIndex, indexer, TODT,OCCURANCEDT, AIRCRAFT, CALLSIGN, AIRPORTS, FDEorMSG, CODE,FAULT)
                    print(listtowrite2)
                    cur2=conn.cursor()
                    cur2.execute("INSERT INTO dbo.[FDE_MSG] VALUES(?,?,?,?,?,?,?,?,?,?)",
                    listtowrite2)
                    conn.commit()
                

    messages[cnt0+1].SaveAs(r'C:\Users\odemirhan\TURISTIK HAVA TASIMACILIK A.S\Gökmen Düzgören - FOE_2019\phyton\db_python\FDE\ACARSmessages'+"/"+str(MailIndex)+".msg")

cnttry=0
while cnttry<15:
    try:
        FDEfolder= outlook.GetDefaultFolder(6).Folders.Item("MAX FDE Messages")
        messages = FDEfolder.Items
        for cnt3 in messages:
            cnt3.Delete()
        if len(messages)!=0:
            cnttry+=1
        else:
            break
        
    except:
        cnttry+=1
