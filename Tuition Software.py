#Paul Classes Tkinter Software Project
 
 
try:
     
     
    #! Program Activation
    import webbrowser
    webbrowser.open('https://i.pinimg.com/originals/60/b4/6a/60b46a38d9cbd6a1cd654fd6d5a679d3.gif')
    from win32com.client import Dispatch

    def speak(text):
        speak=Dispatch(("SAPI.SpVoice"))
        speak.Speak(text)
        
    speak('Activating Software')
    from tkinter import *
    from tkinter import ttk
    from tkinter import messagebox
    import pywhatkit,pandas as pd,string,matplotlib.pyplot as plt,os,clipboard
    from datetime import date
    plt.style.use('dark_background')

    #! Password Entry

    passs=Tk()
    passs.lift()
    passs.title('Password')
    e1=Entry(passs,bg='gray2',fg='blue',show='*',justify='center',font=('Georgia',30),bd=50,insertbackground='blue')
    passs.attributes("-topmost",True)
    e1.grid(row=0,column=0)

    count=0
    access='denied'
    def verify(event):
        
        global count
        global access
        count+=1
        
        if count==3:
            passs.destroy()
            webbrowser.open('https://images-wixmp-ed30a86b8c4ca887773594c2.wixmp.com/f/1f60e133-58e7-4fb7-8b82-3791aeebec0b/dbxx2my-a7d67ae6-909e-4fc6-9db0-dc32aed84eb7.gif?token=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJ1cm46YXBwOiIsImlzcyI6InVybjphcHA6Iiwib2JqIjpbW3sicGF0aCI6IlwvZlwvMWY2MGUxMzMtNThlNy00ZmI3LThiODItMzc5MWFlZWJlYzBiXC9kYnh4Mm15LWE3ZDY3YWU2LTkwOWUtNGZjNi05ZGIwLWRjMzJhZWQ4NGViNy5naWYifV1dLCJhdWQiOlsidXJuOnNlcnZpY2U6ZmlsZS5kb3dubG9hZCJdfQ.UXhaDJEaqDP_uIaDM7Bb-QTglNGoochiidxuXQZ6w90')
            speak('Implementing security punishments , shutting laptop down in 7 seconds')
            print('Shutdown')
            pywhatkit.shutdown(time=7)
            
        elif e1.get()!='sammon12':
            speak('Access Denied '+str(3-count)+' chances left to gain access.')
            e1.delete(0,END)
            l1=Label(passs,text='INVALID PASSWORD ,YOU HAVE '+str(3-count)+' CHANCES TO GET ACCESS',fg='red',bg='darkblue',cursor='pirate',font=('georgia',25))
            l1.grid(row=1,column=0)
            access='denied'
            
        else:
            webbrowser.open('https://i.pinimg.com/originals/bc/55/51/bc5551ac237a9ef4d8e9575662f2e106.gif')
            speak('Access granted')
            passs.destroy()
            access='granted'



    passs.bind('<Return>',verify)

    passs.mainloop()

    #! Current Time Stamp

    def current_date():
        return str(date.today().strftime('%d-%m-%Y'))

    #! Main Window Setting
    speak('Hi Sir, nice to see you again')
    win=Tk()
    win.title('Paul Classes')
    win['background']='gray5'
    win.attributes("-topmost",True)
    win.resizable(0,0)

    #! Main Window Title

    t1=Label(win,text='Paul Classes Software',bg='gray1',fg='green2',font=('georgia',20))
    t1.grid(row=0,column=1,pady=20)

    #! Registration Section:===

    f1=LabelFrame(win,text='Registration Section',bg='gray1',borderwidth=20,fg='white',font=(15))
    f1.grid(row=1,column=0,padx=20)

    #! Registration()

    def Registration():
        

        reg=Tk()
        reg.title('Registration')
        reg['background']='gray27'
        reg.resizable(0,0)
        reg.attributes("-topmost",True)

        t1=Label(reg,text='Registration',bg='gray27',fg='white',font=('georgia',20))
        t1.grid(row=0,column=0,pady=20)

        f1=LabelFrame(reg,text='Register Name',bg='black',fg='white',font=(15),borderwidth=20)
        f1.grid(row=1,column=0,pady=10,padx=20)

        l1=Label(f1,text='Name: ',bg='black',fg='white',font=(10))
        l1.grid(row=0,column=0,pady=10,padx=10,sticky='W')

        e1=Entry(f1,font=(10))
        e1.grid(row=0,column=1,padx=10,sticky='E')

        l2=Label(f1,text='Class: ',bg='black',fg='white',font=(10))
        l2.grid(row=1,column=0,sticky='W',padx=10)

        e2=Entry(f1,font=(10))
        e2.grid(row=1,column=1,sticky='E',padx=10)

        l3=Label(f1,text='School: ',bg='black',fg='white',font=(10))
        l3.grid(row=2,column=0,pady=10,padx=10,sticky='W')

        e3=Entry(f1,font=(10))
        e3.grid(row=2,column=1,sticky='E',padx=10)

        l4=Label(f1,text='Contact Number: ',bg='black',fg='white',font=(10))
        l4.grid(row=3,column=0,sticky='W',padx=10)

        e4=Entry(f1,font=(10))
        e4.grid(row=3,column=1,sticky='E',padx=10)

        l5=Label(f1,text='Remarks: ',bg='black',fg='white',font=(10))
        l5.grid(row=4,column=0,sticky='W',padx=10,pady=10)

        e5=Entry(f1,font=(10))
        e5.grid(row=4,column=1,sticky='E',padx=10)

        def submit1():
            
            try:
                
                data=pd.read_excel('Students List.xlsx')['Name'].tolist()
                if string.capwords(e1.get()) in data:
                    
                    messagebox.showwarning('Name Repetetion','Sir a student by this name has already registered so kindly change the name to avoid any confusion.')
                    return None
                    
            except FileNotFoundError():
                print('File Students List.xlsx not existing')
                pass
            
            if e1.get()=='':
                
                messagebox.showerror('Name please','Sir you forgot to enter a name so kindly enter a name.')
                return None
            
            if e2.get()=='':
                
                messagebox.showerror('Class please','Sir you forgot to enter class ,so kindly enter a class number.')
                return None
            
            if e4.get().isdigit()==False:

                warn=messagebox.showerror('Invalid Contact Number','Sir the contact number entered is not valid so kindly change the number.')
                
                return None
                    
            if os.path.isfile('Students List.xlsx'):

                data=pd.read_excel('Students List.xlsx')
                data=data.append({'Name':string.capwords(e1.get()),'Class':e2.get(),'School':string.capwords(e3.get()),\
                                    'Contact Number':e4.get(),\
                                        'Date Of Joining':current_date(),'Remarks':string.capwords(e5.get())},ignore_index=True)
                data.set_index('Name',inplace=True)
                
                os.remove('Students List.xlsx')
                
                print('Appended Sir')
                
            else:

                data=pd.DataFrame({'Name':[string.capwords(e1.get())],\
                                    'Class':[e2.get()],'School':[string.capwords(e3.get())],\
                                    'Contact Number':[e4.get()],\
                                        'Date Of Joining':current_date(),'Remarks':[string.capwords(e5.get())]})
                data.set_index('Name',inplace=True)

            data.sort_values(by=['Class','Name'],inplace=True)
            data.to_excel('Students List.xlsx')
            reg.destroy()

            que=messagebox.askyesno('Wanna Continue?','Sir do you want to register more names?')

            if que==1:
                Registration()

            else:
                pass
            

        b1=ttk.Button(f1,text='Submit',command=submit1)
        b1.grid(row=5,column=1,pady=20,padx=10,sticky='E')
        
        reg.mainloop()

    b1=ttk.Button(f1,text='Register',command=Registration)
    b1.pack(pady=10)

    #! Record()

    def Record():
        
        ask=messagebox.askyesno("Student's Records?","For students record ,press yes ,else for class records ,press no.")
        if ask==1:

            stu_rec=Tk()
            stu_rec.title('Students Record')
            stu_rec['background']='gray27'
            stu_rec.resizable(0,0)
            stu_rec.attributes("-topmost",True)

            t1=Label(stu_rec,text='Students Record Accessor',bg='gray27',fg='white',font=('georgia',20))
            t1.grid(row=0,column=0,pady=20)
            
            f1=LabelFrame(stu_rec,text='Record Accessor',bg='black',fg='white',font=(15),borderwidth=20)
            f1.grid(row=1,column=0,pady=10,padx=20)
            
            f2=LabelFrame(stu_rec,text='Record',bg='black',fg='white',font=(15),borderwidth=20)
            f2.grid(row=2,column=0,pady=10,padx=20)
            
            l1=Label(f1,text='Sir please select the class of the student: ',bg='black',fg='white',font=(10))
            l1.grid(row=0,column=0,pady=10,padx=10,sticky='W')
            
            data=pd.read_excel('Students List.xlsx')
            
            options=data.Class.unique().tolist()
            
            boxn=ttk.Combobox(f1,values=options,font=(10))
            boxn.grid(row=0,column=1,padx=10,sticky='E')
            
            
            def submit1():
                
            
                l2=Label(f1,text='Sir please select the name of the student: ',bg='black',fg='white',font=(10))
                l2.grid(row=1,column=0,sticky='W',padx=10)
                
                data=pd.read_excel('Students List.xlsx').set_index('Name')
                
                n=StringVar()
                
                if boxn.get().isdigit():
                    cl=int(boxn.get())
                    
                else:
                    cl=string.capwords(boxn.get())
                
                box1=ttk.Combobox(f1,textvariable=n,values=data[data['Class']==cl].index.tolist())
                box1.grid(row=1,column=1,sticky='E',padx=10)
                
                def submit2():
                    
                    l3=Label(f2,text='Name: ',bg='black',fg='yellow',font=(10))
                    l3.grid(row=2,column=0,sticky='W')
                    
                    l4=Label(f2,text=box1.get(),bg='black',fg='white',font=(10))
                    l4.grid(row=2,column=1,sticky='W')
                    
                    l5=Label(f2,text='Class: ',bg='black',fg='yellow',font=(10))
                    l5.grid(row=3,column=0,sticky='W')
                    
                    l6=Label(f2,text=data.loc[box1.get()][0],bg='black',fg='white',font=(10))
                    l6.grid(row=3,column=1,sticky='W')
                    
                    l7=Label(f2,text='School ',bg='black',fg='yellow',font=(10))
                    l7.grid(row=4,column=0,sticky='W')
                    
                    l8=Label(f2,text=data.loc[box1.get()][1],bg='black',fg='white',font=(10))
                    l8.grid(row=4,column=1,sticky='W')
                    
                    l9=Label(f2,text='Contact Information: ',bg='black',fg='yellow',font=(10))
                    l9.grid(row=5,column=0,sticky='W')
                    
                    l10=Label(f2,text=data.loc[box1.get()][2],bg='black',fg='white',font=(10))
                    l10.grid(row=5,column=1,sticky='W')
                    
                    l11=Label(f2,text='Date Of Admission: ',bg='black',fg='yellow',font=(10))
                    l11.grid(row=6,column=0,sticky='W')
                    
                    l12=Label(f2,text=data.loc[box1.get()][3],bg='black',fg='white',font=(10))
                    l12.grid(row=6,column=1,sticky='W')
                    
                    l13=Label(f2,text='Remarks: ',bg='black',fg='yellow',font=(10))
                    l13.grid(row=7,column=0,sticky='W')
                    
                    l14=Label(f2,text=data.loc[box1.get()][4],bg='black',fg='white',font=(10))
                    l14.grid(row=7,column=1,sticky='W')
                    
                    
                    
                    def submit2():
                        stu_rec.destroy()
                        Record()
                    
                    b2=ttk.Button(f1,text='Continue',command=submit2)
                    b2.grid(row=8,column=3)
                    
                    b2=ttk.Button(f1,text='Exit',command=stu_rec.destroy)
                    b2.grid(row=9,column=3)
                    
                b2=ttk.Button(f1,text='Submit',command=submit2)
                b2.grid(row=1,column=2)
                    
            b1=ttk.Button(f1,text='Submit',command=submit1)
            b1.grid(row=0,column=2)
                
        
        
        if ask==0:
            
            stu_rec=Tk()
            stu_rec.title('Class Record')
            stu_rec['background']='gray27'
            stu_rec.resizable(0,0)
            stu_rec.attributes("-topmost",True)

            t1=Label(stu_rec,text='Class Record Accessor',bg='gray27',fg='white',font=('georgia',20))
            t1.grid(row=0,column=0,pady=20)
            
            f2=LabelFrame(stu_rec,text='Record',bg='black',fg='white',font=(15),borderwidth=20)
            f2.grid(row=2,column=0,pady=10,padx=20)
            
            
            f1=LabelFrame(stu_rec,text='',bg='black',fg='white',font=(15),borderwidth=20)
            f1.grid(row=1,column=0,pady=10,padx=20)
            
            l1=Label(f1,text='Sir please enter the class: ',bg='black',fg='white',font=(10))
            l1.grid(row=0,column=0,pady=10,padx=10,sticky='W')
            
            e1=Entry(f1,font=(10))
            e1.grid(row=0,column=1,padx=10,sticky='E')
            
            data=pd.read_excel('Students List.xlsx')
            def submit1():
                data=pd.read_excel('Students List.xlsx')
                
                def cont():
                    
                    stu_rec.destroy()
                    Record()
                    
                b6=ttk.Button(f1,text='Continue',command=cont)
                b6.grid(row=1,column=2,pady=10,sticky='E',padx=5)
                
                b7=ttk.Button(f1,text='Exit',command=stu_rec.destroy)
                b7.grid(row=1,column=1,pady=10,sticky='E')
                
                if e1.get().isdigit():
                    cl=int(e1.get())
                    
                else:
                    cl=string.capwords(e1.get())
                
                data=data[data['Class']==cl]
                roww=0
                columnn=0
                for i in data:
                    l=Label(f2,text=i,bg='black',fg='yellow',font=(10))
                    for k in data[i]:
                        l1=Label(f2,text=k,bg='black',fg='white',font=(10))
                        l1.grid(row=roww+1,column=columnn)
                        roww+=1
                    roww=0
                    l.grid(row=roww,column=columnn)
                    columnn+=1
            
            def generate_excel():
                
                if e1.get().isdigit():
                    cl=int(e1.get())
                    
                else:
                    cl=string.capwords(e1.get())
                
                data=pd.read_excel('Students List.xlsx')
                data=data[data['Class']==cl].set_index('Name')
                data.to_excel('Class '+e1.get()+' records.xlsx')
                clipboard.copy('Class '+e1.get()+' records.xlsx')
                messagebox.showinfo('Excel Generated','Excel has been generated sir and its path has been copied automatically to clipboard sir.')
                
            
            b5=ttk.Button(f2,text='Generate Excel',command=generate_excel)
            b5.grid(row=data.shape[0]+1,column=5)
            
            b1=ttk.Button(f1,text='Submit',command=submit1)
            b1.grid(row=0,column=2)
            

    b2=ttk.Button(f1,text='Records',command=Record)
    b2.pack()

    #! Registration Analysis()

    def analyse():

        anal=Tk()
        anal.title('Analysis')
        anal['background']='gray27'
        anal.resizable(0,0)
        anal.attributes("-topmost",True)
        
        t1=Label(anal,text='Data Analysis',bg='gray27',fg='white',font=('georgia',20))
        t1.grid(row=0,column=0,pady=20)
        
        f1=LabelFrame(anal,text='Analysis',bg='black',fg='white',font=(15),borderwidth=20)
        f1.grid(row=1,column=0)
        
        l1=Label(f1,text='Class wise registration analysis: ',bg='black',fg='white',font=(10))
        l1.grid(row=0,column=0,padx=10,pady=10)
        
        def clas_wis_reg():
            data=pd.read_excel('Students List.xlsx')
            class_lst1=data.Class.unique()
            class_lst=[str(i) for i in class_lst1]
            values=[data[data['Class']==i].shape[0] for i in class_lst]
            print(class_lst,values)
            plt.bar(class_lst,values)
            
            plt.xlabel('Classes')
            plt.ylabel('Class Strength')
            plt.title('Classwise Strength Comparision')
            plt.xticks(class_lst)
            plt.show()
        
        b1=ttk.Button(f1,text='Analyse',command=clas_wis_reg)
        b1.grid(row=0,column=1,padx=10,pady=10)
        
        l2=Label(f1,text='Month wise admission analysis: ',bg='black',fg='white',font=(10))
        l2.grid(row=1,column=0,padx=10,pady=10)
        
        def mon_anal():
            data=pd.read_excel('Students List.xlsx')
            data['Date Of Joining']=pd.to_datetime(data['Date Of Joining'])
            monthly=[]
            for i in range(1,13):
                monthly.append(data[data['Date Of Joining'].dt.month==i].shape[0])
            plt.barh(['January','February','March','April','May','June','July','August'\
                ,'September','October','November','December'],monthly)
            plt.ylabel('Month')
            plt.ylabel('Monthly Admission')
            plt.title('Monthly Admission Analysis')
            plt.show()
        
        b2=ttk.Button(f1,text='Analyse',command=mon_anal)
        b2.grid(row=1,column=1,padx=10,pady=10)
        
    b3=ttk.Button(f1,text='Analysis',command=analyse)
    b3.pack(pady=10)

    #! Attendance Section:==

    f2=LabelFrame(win,text='Attendance Section',bg='gray1',borderwidth=20,fg='white',font=(15))
    f2.grid(row=1,column=1)

    b1=ttk.Button(f2,text='Attendance')
    b1.pack(pady=10)

    b2=ttk.Button(f2,text='Records')
    b2.pack()

    b3=ttk.Button(f2,text='Analysis')
    b3.pack(pady=10)

    #! Finance Section

    f3=LabelFrame(win,text='Finance Section',bg='gray1',borderwidth=20,fg='white',font=(15))
    f3.grid(row=1,column=2,padx=20)

    def deposit():
        
        dep=Tk()
        dep.title('Deposits')
        dep['background']='gray27'
        dep.resizable(0,0)
        dep.attributes("-topmost",True)
        
        t1=Label(dep,text='Deposits',bg='gray27',fg='white',font=('georgia',20))
        t1.grid(row=0,column=0,pady=20)
        
        f1=LabelFrame(dep,text='Deposit Entry',bg='black',fg='white',font=(15),borderwidth=20)
        f1.grid(row=1,column=0,pady=10)
        
        l1=Label(f1,text='Class: ',bg='black',fg='white',font=(10))
        l1.grid(row=0,column=0)
        
        data=pd.read_excel('Students List.xlsx')
        
        options=data.Class.unique().tolist()
        
        box4=ttk.Combobox(f1,values=options,font=(10))
        box4.grid(row=0,column=1,padx=10,pady=10,sticky='W')
        
        def submit():
            
            l2=Label(f1,text='Name: ',bg='black',fg='white',font=(10))
            l2.grid(row=1,column=0)
            
            if box4.get().isdigit():
                stu_class=int(box4.get())
                
            else:
                stu_class=string.capwords(box4.get())
            
            data=pd.read_excel('Students List.xlsx')
            options=data[data['Class']==stu_class]['Name'].tolist()
            
            n=StringVar()
            box1=ttk.Combobox(f1,textvariable=n,values=options)
            box1.grid(row=1,column=1,sticky='W',padx=10)
            
            def submit2():
                
                stu_name=box1.get()
                
                if stu_name not in options:
                    
                    messagebox.showwarning('Invalid Name','Sir we have no student by this name so please select a correct name.')
                    box1.set('')
                    return None
                
                l3=Label(f1,text='Fees Deposited: ',bg='black',fg='white',font=(10))
                l3.grid(row=2,column=0)
                
                e2=Entry(f1,font=(10))
                e2.grid(row=2,column=1,padx=10,pady=10,sticky='W')
                
                def submit3():
                    
                    if e2.get().isdigit()==False:
                        
                        messagebox.showwarning('Invalid Fees','Sir please enter a valid fee amount.')
                        e2.delete(0,END)
                        return None
                        
                
                    if os.path.isfile('Fee Deposits.xlsx')==False:
                        
                        dataf=pd.DataFrame({'Name':[string.capwords(stu_name)],'Class':[box4.get()],'Fee Deposited':float(e2.get()),'Date':[current_date()]})
                        dataf.set_index('Name',inplace=True)
                        dataf.to_excel('Fee Deposits.xlsx')
                        
                    else:
                        
                        dataf=pd.read_excel('Fee Deposits.xlsx')
                        dataf=dataf.append({'Name':string.capwords(stu_name),'Class':box4.get(),'Fee Deposited':float(e2.get()),\
                                            'Date':current_date()},ignore_index=True)
                        dataf.set_index('Name',inplace=True)
                        
                        os.remove('Fee Deposits.xlsx')
                        dataf.to_excel('Fee Deposits.xlsx')
                        print('Appended Sir')
                        
                        l4=Label(f1,text='Deposit Recorded Successfully.',bg='black',fg='blue',font=(10))
                        l4.grid(row=3,column=0)
                        
                    ask=messagebox.askyesno("Continue?","Fee record of "+string.capwords(box1.get())+" stored successfully ,do you want to continue?")
                    
                    if ask==1:
                        
                        dep.destroy()
                        deposit()
                        
                    else:
                        
                        dep.destroy()
                    
                b3=ttk.Button(f1,text='Submit',command=submit3)
                b3.grid(row=2,column=2)
                        
            b2=ttk.Button(f1,text='Submit',command=submit2)
            b2.grid(row=1,column=2)
                        
        b1=ttk.Button(f1,text='Submit',command=submit)
        b1.grid(row=0,column=2)

    b1=ttk.Button(f3,text='Deposit',command=deposit)
    b1.pack(pady=10)

    def rec():
        
        recd=Tk()
        recd.title('Records')
        recd['background']='gray27'
        recd.resizable(0,0)
        recd.attributes("-topmost",True)
        
        t1=Label(recd,text='Fee Records',bg='gray27',fg='white',font=('georgia',20))
        t1.grid(row=0,column=0,pady=20)
        
        f1=LabelFrame(recd,text='Records',bg='black',fg='white',font=(15),borderwidth=20)
        f1.grid(row=1,column=0)
        
        l1=Label(f1,text='Class: ',bg='black',fg='white',font=(10))
        l1.grid(row=0,column=0,padx=10,pady=10)
        
        data=pd.read_excel('Fee Deposits.xlsx').set_index('Name')
        
        box1=ttk.Combobox(f1,values=['All']+data.Class.unique().tolist())
        box1.grid(row=0,column=1,sticky='W',padx=10)
        
        def submit():
            data=pd.read_excel('Fee Deposits.xlsx')
            ls=['All']+data.Class.unique().tolist()
            
            cl=box1.get()
            
            if cl.isdigit():
                cl=int(cl)
            
            if cl not in ls:
                
                messagebox.showwarning('Invalid Class','Sir no class like this exits so kindly choose an existing class.')
                box1.set('')
                return None
            
            stu_class=cl
            
            f2=LabelFrame(recd,text='Data',bg='black',fg='white',font=(15),borderwidth=20)
            f2.grid(row=2,column=0)
            
            if cl=='All':
                
                lst1=ls
                
            else:
                
                lst1=[cl]
            
            Name=data[data['Class'].isin(lst1)]['Name'].tolist()
            Class=data[data['Class'].isin(lst1)]['Class'].tolist()
            Fee=data[data['Class'].isin(lst1)]['Fee Deposited'].tolist()
            Date=data[data['Class'].isin(lst1)]['Date'].tolist()
            
            #Putting Column name
            
            l1=Label(f2,text='Name',bg='black',fg='gold',font=(10))
            l1.grid(row=0,column=0)
            
            l2=Label(f2,text='Class',bg='black',fg='gold',font=(10))
            l2.grid(row=0,column=1)
            
            l3=Label(f2,text='Fee History',bg='black',fg='gold',font=(10))
            l3.grid(row=0,column=2)
            
            l4=Label(f2,text='Date',bg='black',fg='gold',font=(10))
            l4.grid(row=0,column=3)
            
            #Making table
            count=0
            
            for i in Name:
                
                count+=1    
                l=Label(f2,text=i,bg='black',fg='yellow',font=(10))
                l.grid(row=count,column=0)
                
            count=0
                
            for i in  Class:
                
                count+=1
                l=Label(f2,text=i,bg='black',fg='white',font=(10))
                l.grid(row=count,column=1)
                
            count=0
                
            for i in  Fee:
                
                count+=1
                l=Label(f2,text=i,bg='black',fg='white',font=(10))
                l.grid(row=count,column=2)
                
            count=0
                
            for i in  Date:
                
                count+=1
                l=Label(f2,text=i,bg='black',fg='white',font=(10))
                l.grid(row=count,column=3)
                
            def excel():
                data=pd.read_excel('Fee Deposits.xlsx')
                data=data[data['Class'].isin(lst1)].set_index('Name')
                
                if cl!='All':
                    data.to_excel('Fee Depositors of Class '+str(cl)+'.xlsx')
                    clipboard.copy('Fee Depositors of Class '+str(cl)+'.xlsx')
                    messagebox.showinfo('Excel Generated','Excel file for this data generated by name : '+'Fee Depositors of Class '+str(cl)+'.xlsx has been copied to clipboard.')
                    
                if cl=='All':
                    data.to_excel('Fee Depositors.xlsx')
                    clipboard.copy('Fee Depositors.xlsx')
                    messagebox.showinfo('Excel Generated','Excel file for this data generated by name : Fee Depositors.xlsx has been copied to clipboard.')

                
            b2=ttk.Button(f2,text='Generate Excel',command=excel)
            b2.grid(row=count+1,column=3,padx=10,pady=10)
            
            b3=ttk.Button(f2,text='Exit',command=recd.destroy)
            b3.grid(row=count+2,column=3,pady=10)
        
        b1=ttk.Button(f1,text='Submit',command=submit)
        b1.grid(row=0,column=2)

    b2=ttk.Button(f3,text='Records',command=rec)
    b2.pack()
    
    def receipt():
        
        rec=Tk()
        rec.title('Records')
        rec['background']='gray27'
        rec.resizable(0,0)
        rec.attributes("-topmost",True)
        
        t1=Label(rec,text='Receipt',bg='gray27',fg='white',font=('georgia',20))
        t1.grid(row=0,column=0,pady=20)
        
        f1=LabelFrame(rec,text='Entry',bg='black',fg='white',font=(15),borderwidth=20)
        f1.grid(row=1,column=0)
        
        l1=Label(f1,text='Class: ',bg='black',fg='white',font=(10))
        l1.grid(row=0,column=0,padx=10,pady=10)
        
        data=pd.read_excel('Fee Deposits.xlsx')
        
        box1=ttk.Combobox(f1,values=data.Class.unique().tolist())
        box1.grid(row=0,column=1,padx=10)
        
        def submit():
            
            stu_class=box1.get()
            
            if stu_class.isdigit():
                stu_class=int(stu_class)
                
            data=pd.read_excel('Fee Deposits.xlsx')
            stu=data[data['Class']==stu_class]['Name'].tolist()
            
            l2=Label(f1,text='Name: ',bg='black',fg='white',font=(10))
            l2.grid(row=1,column=0,padx=10,pady=10)
            
            box2=ttk.Combobox(f1,values=stu)
            box2.grid(row=1,column=1)
            
            def submit1():
                data=pd.read_excel('Fee Deposits.xlsx')
                name=box2.get()
                f2=LabelFrame(rec,text='Receipt',bg='ghost white',fg='blue',font=(15),borderwidth=20)
                f2.grid(row=2,column=0,pady=15)
                
                t1=Label(f2,text='Paul Classes',bg='ghost white',fg='blue',font=('Georgia',20))
                t1.grid(row=0,column=1)
                
                l1=Label(f2,text=current_date(),bg='ghost white',fg='green2',font=(10))
                l1.grid(row=1,column=1)
                
                t2=Label(f2,text=name,bg='ghost white',fg='red2',font=('Georgia',13))
                t2.grid(row=2,column=1,pady=10)
                
                dates=data[data['Name']==name]['Date'].tolist()
                fees=data[data['Name']==name]['Fee Deposited'].tolist()
                
                ind=0
                
                for i in dates:
                    
                    l=Label(f2,text=dates[ind]+'====> Rupees '+str(fees[ind]),bg='ghost white',fg='purple',font=(10))
                    l.grid(row=ind+3,column=0,columnspan=2,sticky='W')
                    
                    ind+=1
                    
                l=Label(f2,text='Total Fee Paid: '+str(sum(fees)),bg='ghost white',fg='red',font=(10))
                l.grid(row=ind+3,column=0,columnspan=2,pady=10,sticky='W')
            
            b2=ttk.Button(f1,text='Submit',command=submit1)
            b2.grid(row=1,column=2)
        
        b1=ttk.Button(f1,text='Submit',command=submit)
        b1.grid(row=0,column=2,padx=10)

    b3=ttk.Button(f3,text='Receipt',command=receipt)
    b3.pack(pady=10)

    f4=LabelFrame(win,text='Useful Softwares',bg='gray1',borderwidth=20,fg='white',font=(15))
    f4.grid(row=2,column=1,pady=30)

    #! Bible Section:==

    def bible():
        webbrowser.open('https://www.worldhistory.biz/download567/The_Orthodox_Study_Bible_-_St.pdf')

    b1=ttk.Button(f4,text='Bible',command=bible)
    b1.pack(pady=10)

    #! Whatsapp Section

    def whatsapp():
        webbrowser.open('https://web.whatsapp.com/')
        
    b2=ttk.Button(f4,text='Whatsapp',command=whatsapp)
    b2.pack()

    #! Google Section:==

    def google():
        
        
        g=Tk()
        g.title('Google Search')
        g['background']='gray27'
        g.resizable(0,0)
        g.attributes("-topmost",True)
        
        t1=Label(g,text='Google Assistance',bg='gray27',fg='white',font=('georgia',20))
        t1.grid(row=0,column=0,pady=20)
        
        f1=LabelFrame(g,text='Google',bg='black',fg='white',font=(15),borderwidth=20)
        f1.grid(row=1,column=0,pady=10)
        
        l1=Label(f1,text='Please enter what you want to search in Google: ',bg='black',fg='white',font=(10))
        l1.grid(row=0,column=0,padx=10,pady=10)
        
        e1=Entry(f1,font=(10))
        e1.grid(row=0,column=1,padx=10,pady=10)
        
        def searching(event):
            pywhatkit.search(e1.get())
        
        g.bind('<Return>',searching)
        
        f2=LabelFrame(g,text='Google Workspace',bg='black',fg='white',font=(15),borderwidth=20)
        f2.grid(row=2,column=0,pady=10)
        
        def goog():
            webbrowser.open('www.google.com')
        
        b1=ttk.Button(f2,text='Google',command=goog)
        b1.grid(row=0,column=0,padx=10,pady=10)
        
        def gmail():
            webbrowser.open('www.gmail.com')
        
        b2=ttk.Button(f2,text='G Mail',command=gmail)
        b2.grid(row=0,column=1,padx=10,pady=10)
        
        def google_meet():
            webbrowser.open('https://meet.google.com/')
        
        b3=ttk.Button(f2,text='Google Meet',command=google_meet)
        b3.grid(row=0,column=2,padx=10,pady=10)
        
        def google_sheets():
            webbrowser.open('https://www.google.com/sheets/about/')
        
        b4=ttk.Button(f2,text='Google Sheets',command=google_sheets)
        b4.grid(row=0,column=3,padx=10,pady=10)
        
        def google_photos():
            webbrowser.open('https://photos.google.com/')
        
        b5=ttk.Button(f2,text='Google Photos',command=google_photos)
        b5.grid(row=1,column=0,padx=10,pady=10)
        
        def google_slides():
            webbrowser.open('https://www.google.com/slides/about/')
        
        b6=ttk.Button(f2,text='Google Slides',command=google_slides)
        b6.grid(row=1,column=1,padx=10,pady=10)
        
        def google_maps():
            webbrowser.open('https://www.google.com/maps')

        b7=ttk.Button(f2,text='Google Maps',command=google_maps)
        b7.grid(row=1,column=2,padx=10,pady=10)
        
        def google_calendar():
            webbrowser.open('https://www.google.com/calendar/about/')
        
        b8=ttk.Button(f2,text='Google Calendar',command=google_calendar)
        b8.grid(row=1,column=3,padx=10,pady=10)
        
        def google_drive():
            webbrowser.open('https://www.google.com/intl/en_ca/drive/download/')

        b9=ttk.Button(f2,text='Google Drive',command=google_drive)
        b9.grid(row=2,column=0,padx=10,pady=10)
        
        def google_translate():
            webbrowser.open('https://translate.google.co.in/')
        
        b10=ttk.Button(f2,text='Google Translate',command=google_translate)
        b10.grid(row=2,column=1,padx=10,pady=10)
        
        def google_classroom():
            webbrowser.open('https://classroom.google.com/h')
        
        b11=ttk.Button(f2,text='Google Classroom',command=google_classroom)
        b11.grid(row=2,column=2,padx=10,pady=10)
        
        def google_docs():
            webbrowser.open('https://docs.google.com/document/u/0/')
        
        b12=ttk.Button(f2,text='Google Docs',command=google_docs)
        b12.grid(row=2,column=3,padx=10,pady=10)

    b1=ttk.Button(f4,text='Google',command=google)
    b1.pack(pady=10)

    #! YouTube Section:==

    def you_tube():
        
        y=Tk()
        y.title('Google Search')
        y['background']='gray27'
        y.resizable(0,0)
        y.attributes("-topmost",True)
        
        t1=Label(y,text='You Tube Assistance',bg='gray27',fg='white',font=('georgia',20))
        t1.grid(row=0,column=0,pady=20)
        
        f1=LabelFrame(y,text='Youtube',bg='black',fg='white',font=(15),borderwidth=20)
        f1.grid(row=1,column=0,pady=10)
        
        l1=Label(f1,text='Please enter what you want to search in You Tube: ',bg='black',fg='white',font=(10))
        l1.grid(row=0,column=0,padx=10,pady=10)
        
        e1=Entry(f1,font=(10))
        e1.grid(row=0,column=1,padx=10,pady=10)
        
        def searching():
            pywhatkit.playonyt(e1.get())
        
        b1=ttk.Button(f1,text='Search',command=searching)
        b1.grid(row=0,column=2,padx=10,pady=10)
        
        def open_yt():
            webbrowser.open('www.youtube.com')
            
        b2=ttk.Button(f1,text='You Tube',command=open_yt)
        b2.grid(row=1,column=0,sticky='E',pady=20) 

    b2=ttk.Button(f4,text='You Tube',command=you_tube)
    b2.pack()

    #! GeoGebra Section:==

    def geo_gebra():
        webbrowser.open('https://www.geogebra.org/?lang=en')

    b3=ttk.Button(f4,text='Geo Gebra',command=geo_gebra)
    b3.pack(pady=10)

    #! Spotify Section:==

    def spotify():
        webbrowser.open('https://open.spotify.com/')

    b4=ttk.Button(f4,text='Spotify',command=spotify)
    b4.pack()

    #! Shutdown Section:==

    def shutdown():
        msg=messagebox.askquestion('Shut Down?','Sir are you sure to shut down?')
        if msg=='yes':
            pywhatkit.shutdown(time=5)
        else:
            pass

    b4=ttk.Button(win,text='Shutdown',command=shutdown)
    b4.grid(row=3,column=1,pady=10)

    #! Exit Section:==

    def exit():
        speak('Terminating program ,thanks for giving me an oppurtunity to serve you sir.')
        win.destroy()

    ex1=ttk.Button(win,text='Exit',command=exit)
    ex1.grid(row=4,column=1,pady=10)

    if access=='denied':
        win.destroy()

    win.mainloop()

except Exception as e:
    
    speak("Very sorry to say that an unexpected error has taken place in the software which is being described as \
        "+e+" so kindly wait till this program starts it recovery system")
    webbrowser.open("https://i.pinimg.com/originals/4c/22/18/4c2218f5cc96ba76c0e590cd1dadb1bc.gif")
    os._exit(0)
