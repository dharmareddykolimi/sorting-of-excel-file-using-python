from datetime import datetime
import pandas as pd
import xlsxwriter
from tkinter import *
class MyWindow:
    def __init__(self, win):
        self.lbl1=Label(win, text='location of file')
        self.lbl2=Label(win, text='location to save ')
        self.lbl3=Label(win, text='loca of file saved')
        self.t1=Entry(bd=3)
        self.t2=Entry()
        self.t3=Entry()
        self.btn1 = Button(win, text='split')
        self.lbl1.place(x=100, y=50)
        self.t1.place(x=200, y=50)
        self.lbl2.place(x=100, y=100)
        self.t2.place(x=200, y=100)
        self.b1=Button(win, text='split', command=self.split)
        self.b1.place(x=100, y=150)
        self.lbl3.place(x=100, y=200)
        self.t3.place(x=200, y=200)
    def split(self):
        self.t3.delete(0, 'end')
        location = str(self.t1.get())
        save = str(self.t2.get())
        df = pd.read_excel(location,header=None)
        p=1
        CSEA = ["CSE-A"]
        CSEB = ["CSE-B"]
        CSEC = ["CSE-C"]
        CSED = ["CSE-D"]
        dat= df[0][0]
        rows=len(df)
        columns=len(df.columns)
        x = range(1, rows)
        y = range(0,columns)
        f=0
        for i in x:
            for j in y:
                if df[j][i]=="":
                    print('')
                else:
                    str1=str(df[j][i])
                    if str1[0]=='1':
                         if str1[1]=='7' or str1[1]=='8' or str1[1]=='6':
                                if (str1[1]=='7' and ((str1[8]=='0'or str1[8]=='1'or str1[8]=='2'or str1[8]=='3'or str1[8]=='4'or str1[8]=='5') or (str1[8]=='6' and str1[9]=='0'))) or str1[1]=='6' :

                                    if str1 in CSEA:
                                        iam = 'KDR'
                                    else:
                                        CSEA.append(str1)

                                    CSEA.sort()

                                if (str1[1]=='7' and ((str1[8]=='6'or str1[8]=='7'or str1[8]=='8'or str1[8]=='9'or str1[8]=='A'or str1[8]=='B')or ( str1[8]=='C' and str1[9]=='0'))) or (( str1[1]=='8' and str1[8]=='0') and (str1[9]=='1' or str1[9]=='2' or str1[9]=='3' or str1[9]=='4' or str1[9]=='5' or str1[9]=='6' or str1[9]=='8' or str1[9]=='9') or (str1[1]=='8' and str1[8]=='1' and str1[9]=='0')):

                                    if str1 in CSEB:
                                        iam = 'KDR'
                                    else:
                                        CSEB.append(str1)

                                    CSEB.sort()
                                if (str1[1]=='7' and ((str1[8]=='C'or str1[8]=='D'or str1[8]=='E'or str1[8]=='F'or str1[8]=='G'or str1[8]=='H') or ( str1[8]=='J' and str1[9]=='0')))  or ((str1[1]=='8' and str1[8]=='1') and (str1[9]=='1'or str1[9]=='2'or str1[9]=='3'or str1[9]=='5'or str1[9]=='6'or str1[9]=='7'or str1[9]=='8'or str1[9]=='9')) or ((str1[1]=='8' and str1[8]=='2') and (str1[9]=='0'or str1[9]=='1'or str1[9]=='2'or str1[9]=='3')) :

                                    if str1 in CSEC:
                                        iam = 'KDR'
                                    else:
                                        CSEC.append(str1)

                                    CSEC.sort()
                                if (str1[1]=='7' and ((str1[8]=='J'or str1[8]=='K'or str1[8]=='O'or str1[8]=='L'or str1[8]=='M'or str1[8]=='N'or str1[8]=='P') or ( str1[8]=='Q' and str1[9]=='0'))) or ((str1[1]=='8' and str1[8]=='2') and (str1[9]=='4'or str1[9]=='5')) :

                                    if str1 in CSED:
                                        iam = 'KDR'
                                    else:
                                        CSED.append(str1)

                                    CSED.sort()
        workbook = xlsxwriter.Workbook(save)
        worksheet = workbook.add_worksheet()
        lena=len(CSEA)
        p=1
        aa=0
        for aa in range(lena):
            A='A'
            Ac=A +str(p) 
            worksheet.write(Ac,CSEA[aa])
            p=p+1

        lenb=len(CSEB)
        q=1
        ab=0
        for ab in range(lenb):
            A='B'
            Ac=A +str(q) 
            worksheet.write(Ac,CSEB[ab])
            q=q+1

        lenc=len(CSEC)
        r=1
        ac=0
        for ac in range(lenc):
            A='C'
            Ac=A +str(r) 
            worksheet.write(Ac,CSEC[ac])
            r=r+1

        lend=len(CSED)
        s=1
        ad=0
        for ad in range(lend):
            A='D'
            Ac=A +str(s) 
            worksheet.write(Ac,CSED[ad])
            s=s+1
        worksheet.write('E1',str(dat))
        workbook.close()
        result=save
        self.t3.insert(END, str(result))

window=Tk()
mywin=MyWindow(window)
window.title('SPLIT THE FILE BY KDR')
window.geometry("600x600+20+20")
window.mainloop()

