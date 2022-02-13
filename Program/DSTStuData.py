from tkinter import *
from tkinter import filedialog
from tkinter import messagebox as msg

import sys
#from numpy import imag
sys.path.insert(1,'DST\Program')
import DSTStuData_BackEnd as BE
from PIL import ImageTk, Image
import os
if BE.DEBUG:
    DSTFolderPath = ''
else:
    DSTFolderPath = BE.DSTFolderPath 
    
SUBJECTS = BE.SUBJECTS
TEACHERS = BE.TEACHERS


fontStyle = 'Times'

#Add a cancel button to not take any value

def main():
    root = Tk()
    HEIGHT = 200
    WIDTH = 600
    root.geometry('600x200+50+50')
    root.iconbitmap(DSTFolderPath + 'Program\Images\Logo.ico')
    root.title('DST Data Program')

    DownArrowImg = ImageTk.PhotoImage(Image.open(DSTFolderPath + 'Program\Images\DownArrow.png'))

    
    def opendial():
        filename = filedialog.askopenfilename(initialdir= os.path.relpath(__file__), title='Select the file', filetypes= [("Excel Sheet","*.xlsx")])
        return filename

    def IncompErrorMsg():
        msg.showinfo('Notification', 'Some fields were not entered\nNo changes made')

    def AddMarks():
        marksFilePath = StringVar()
        marksFilePath.set('')
        attendanceFilePath = StringVar()
        attendanceFilePath.set('')
        SubjectVar = StringVar()
        SubjectVar.set('Select Subject')
        level = Toplevel()
        level.geometry('500x200+400+400')

        level.title('DST Student')

        Title = Label(level, text= "Add Subject Marks")
        Title.config(font=(fontStyle, 20))
        Title.place(relx=0.3, rely=0.01)


        MarksLabel = Label(level, text= "Select the Student Marks file ")
        MarksLabel.place(relx=0.01, rely=0.2)
        MarksLabel.config(font=(fontStyle, 12))

        def MarksFilePath():
            marksFilePath.set(opendial())
            if(marksFilePath.get() == ''):
                return ('')
            x = marksFilePath.get().split('/')[-1]
            MarksButton.config(text= x)

        MarksButton = Button(level, text="Select file", command=lambda: MarksFilePath())
        MarksButton.place(relx=0.5, rely=0.2)
        MarksButton.config(font=(fontStyle, 12))


        AttendanceLabel = Label(level, text= "Select the Student Attendance file ")
        AttendanceLabel.place(relx=0.01, rely=0.4)
        AttendanceLabel.config(font=(fontStyle, 12))

        def AttendanceFilePath():
            attendanceFilePath.set(opendial())
            if(attendanceFilePath.get() == ''):
                return ('')
            x = attendanceFilePath.get().split('/')[-1]
            AttendanceButton.config(text= x)

        AttendanceButton = Button(level, text="Select file", command=lambda: AttendanceFilePath())
        AttendanceButton.config(font=(fontStyle, 12))
        AttendanceButton.place(relx=0.5, rely=0.4)


        SubjectLabel = Label(level, text= "Subject ")
        SubjectLabel.place(relx=0.01, rely=0.6)
        SubjectLabel.config(font=(fontStyle, 12))

        
        SubjectInput = OptionMenu(level, SubjectVar, *SUBJECTS)
        SubjectInput.place(relx=0.2, rely=0.6)
        SubjectInput.config(indicatoron=0, font=(fontStyle, 12), compound='right', image= DownArrowImg)

        global flag
        flag = True

        def Execute():
            level.quit()
            level.destroy()

        ExecuteButton = Button(level, text="Execute", command=Execute)
        ExecuteButton.config(font=(fontStyle, 12))
        ExecuteButton.place(relx=0.8, rely=0.8)

        def Cancel():
            global flag
            flag = False
            level.quit()
            level.destroy()

        CanceleButton = Button(level, text="Cancel", command=Cancel)
        CanceleButton.config(font=(fontStyle, 12))
        CanceleButton.place(relx=0.66, rely=0.8)


        level.mainloop()

        if(SubjectVar.get() == 'Select Subject'):
            SubjectVar.set('')
        if not flag:
            return('','','')
        
        subTeacher = ''
        for i in range(len(SUBJECTS)):
            if(SUBJECTS[i] == SubjectVar.get()):
                subTeacher = TEACHERS[i]
        return(marksFilePath.get(), attendanceFilePath.get(), SubjectVar.get(), subTeacher)

    def AddStudent():
        detailsFilePath = StringVar()
        detailsFilePath.set('')
        level = Toplevel()
        level.title('DST Student')
        level.geometry('375x150+400+400')

        Title = Label(level, text= "Add a Student")
        Title.config(font=(fontStyle, 20))
        Title.place(relx=0.3, rely=0.01)

        DetailsLabel = Label(level, text= "Select the Student Details file ")
        DetailsLabel.place(relx=0.01, rely=0.35)
        DetailsLabel.config(font=(fontStyle, 12))

        def DetailsFilePath():
            detailsFilePath.set(opendial())
            if(detailsFilePath.get() == ''):
                return ('')
            x = detailsFilePath.get().split('/')[-1]
            DetailsButton.config(text= x)

        DetailsButton = Button(level, text="Select file", font=2, command=lambda: DetailsFilePath())
        DetailsButton.place(relx=0.5, rely=0.35)
        DetailsButton.config(font=(fontStyle, 12))

        global flag
        flag = True

        def Execute():
            level.quit()
            level.destroy()

        ExecuteButton = Button(level, text="Execute", font=10, command=lambda: Execute())
        ExecuteButton.config(font=(fontStyle, 12))
        ExecuteButton.place(relx=0.8, rely=0.75)

        def Cancel():
            global flag
            flag = False
            level.quit()
            level.destroy()

        CanceleButton = Button(level, text="Cancel", command=Cancel)
        CanceleButton.config(font=(fontStyle, 12))
        CanceleButton.place(relx=0.61, rely=0.75)

        level.mainloop()

        if not flag:
            return('')
        return(detailsFilePath.get())

    def AddFees():
        feesFilePath = StringVar()
        feesFilePath.set('')
        level = Toplevel()

        level.title("DST Student")
        level.geometry('400x175+400+400')


        Title = Label(level, text= "Add a Year's Fees")
        Title.config(font=(fontStyle, 20))
        Title.place(relx=0.3, rely=0.01)


        FeesLabel = Label(level, text= "Select the Students Fees file ")
        FeesLabel.place(relx=0.01, rely=0.25)
        FeesLabel.config(font=(fontStyle, 12))

        def FeesFilePath():
            feesFilePath.set(opendial()) 
            if (feesFilePath.get() == ''):
                return ('')
            x = feesFilePath.get().split('/')[-1]
            FeesButton.config(text= x)

        FeesButton = Button(level, text="Select file", font=2, command=lambda: FeesFilePath())
        FeesButton.place(relx=0.5, rely=0.25)
        FeesButton.config(font=(fontStyle, 12))


        YearLabel = Label(level, text= "Year ")
        YearLabel.place(relx=0.01, rely=0.5)
        YearLabel.config(font=(fontStyle, 12))

        YearInput = Entry(level)
        YearInput.place(relx=0.2, rely=0.5)
        YearInput.config(font=(fontStyle, 12))
        global year
        year = YearInput.get()

        global flag
        flag = True

        def Execute():
            global year
            year = YearInput.get()
            level.quit()
            level.destroy()

        ExecuteButton = Button(level, text="Execute", font=10, command=Execute)
        ExecuteButton.config(font=(fontStyle, 12))
        ExecuteButton.place(relx=0.8, rely=0.75)

        def Cancel():
            global flag
            flag = False
            level.quit()
            level.destroy()

        CanceleButton = Button(level, text="Cancel", command=Cancel)
        CanceleButton.config(font=(fontStyle, 12))
        CanceleButton.place(relx=0.61, rely=0.75)

        level.mainloop()

        if not flag:
            return ('','')

        return(feesFilePath.get(), year)

    def DeleteFees():
        level = Toplevel()

        level.title('DST Student')
        level.geometry('375x125+400+400')

        Title = Label(level, text= "Delete Fees Row")
        Title.config(font=(fontStyle, 20))
        Title.place(relx=0.3, rely=0.01)

        YearLabel = Label(level, text= "Year ")
        YearLabel.place(relx=0.01, rely=0.35)
        YearLabel.config(font=(fontStyle, 12))

        YearInput = Entry(level)
        YearInput.place(relx=0.15, rely=0.35)
        YearInput.config(font=(fontStyle, 12))

        global flag
        flag = True

        def Execute():
            global year
            year = YearInput.get()
            level.quit()
            level.destroy()

        ExecuteButton = Button(level, text="Execute", font=10, command=Execute)
        ExecuteButton.config(font=(fontStyle, 12))
        ExecuteButton.place(relx=0.8, rely=0.7)

        def Cancel():
            global flag
            flag = False
            level.quit()
            level.destroy()

        CanceleButton = Button(level, text="Cancel", command=Cancel)
        CanceleButton.config(font=(fontStyle, 12))
        CanceleButton.place(relx=0.61, rely=0.7)

        level.mainloop()

        if not flag:
            return(year)
        return(year)

    def DeleteMarks():
        level = Toplevel()

        level.title('DST Student')
        level.geometry('375x150+400+400')
        Title = Label(level, text= "Delete Subject Marks")
        Title.config(font=(fontStyle, 20))
        Title.place(relx=0.3, rely=0.01)

        SubjectLabel = Label(level, text= "Subject ")
        SubjectLabel.place(relx=0.01, rely=0.33)
        SubjectLabel.config(font=(fontStyle, 12))

        SubjectVar = StringVar()
        SubjectVar.set('Select Subject')
        SubjectInput = OptionMenu(level, SubjectVar, *SUBJECTS)
        SubjectInput.place(relx=0.15, rely=0.33)
        SubjectInput.config(indicatoron=0, font=(fontStyle, 12), compound= 'right', image= DownArrowImg)

        global flag
        flag = True

        def Execute():
            level.quit()
            level.destroy()

        ExecuteButton = Button(level, text="Execute", font=10, command=Execute)
        ExecuteButton.config(font=(fontStyle, 12))
        ExecuteButton.place(relx=0.8, rely=0.75)

        def Cancel():
            global flag
            flag = False
            level.quit()
            level.destroy()

        CanceleButton = Button(level, text="Cancel", command=Cancel)
        CanceleButton.config(font=(fontStyle, 12))
        CanceleButton.place(relx=0.61, rely=0.75)

        level.mainloop()
        if(SubjectVar.get() == 'Select Subject'):
            SubjectVar.set('')
        if not flag:
            return('')
        return(SubjectVar.get())

    def DeleteStudent():
        level = Toplevel()
        level.title('DST Student')
        level.geometry('375x150+400+400')

        Title = Label(level, text= "Remove A Student")
        Title.config(font=(fontStyle, 20))
        Title.place(relx=0.25, rely=0.01)

        RollLabel = Label(level, text= "Roll no. ")
        RollLabel.place(relx=0.01, rely=0.3)
        RollLabel.config(font=(fontStyle, 12))
        RollInput = Entry(level)
        RollInput.place(relx=0.15, rely=0.33)
        RollInput.config(font=(fontStyle, 12))
        
        global flag
        flag = True
        global roll

        def Execute():
            global roll
            roll = RollInput.get()
            level.quit()
            level.destroy()

        ExecuteButton = Button(level, text="Execute", font=10, command=Execute)
        ExecuteButton.config(font=(fontStyle, 12))
        ExecuteButton.place(relx=0.8, rely=0.75)

        def Cancel():
            global flag
            flag = False
            level.quit()
            level.destroy()

        CanceleButton = Button(level, text="Cancel", command=Cancel)
        CanceleButton.config(font=(fontStyle, 12))
        CanceleButton.place(relx=0.61, rely=0.75)

        level.mainloop()
        if not flag:
            return('')
        return(roll)
#Main

    LogoImg = ImageTk.PhotoImage(Image.open(DSTFolderPath + 'Program\Images\Logo1.png'))
    ImageTitle = Label(image = LogoImg)
    ImageTitle.place(relx=0.17, rely=0.03)
    Title = Label(root, text= "Divine School\nof Theology")
    Title.config(font=("Courier", 30))
    Title.place(relx=0.3, rely=0.01)


    def AddDataFunc(x):
        if(x != 'Add Data'):
            if(x == 'Student'):
                p = AddStudent()
                if(p == ''):
                    IncompErrorMsg()
                else:
                    msg.showinfo('Notification', BE.AddStudents(p))
            elif(x == 'Marks of a Subject'):
                p = AddMarks()
                if('' in p):
                    IncompErrorMsg()
                else:
                    msg.showinfo('Notification', BE.AddMarks(p[0], p[1], p[2], p[3]))
            elif(x == 'Fees of a Year'):
                p = AddFees()
                if('' in p):
                    IncompErrorMsg()
                else:
                    msg.showinfo('Notification', BE.AddFees(p[0], p[1]))
            clickedAdd.set("Add Data")

    clickedAdd = StringVar()
    clickedAdd.set('Add Data')
    AddDataList = ['Student', 'Marks of a Subject', 'Fees of a Year']
    AddDataOptions = OptionMenu(root, clickedAdd, *AddDataList, command= AddDataFunc)
    AddDataOptions.place(relx = 0.1, rely = 0.5)
    AddDataOptions.config(indicatoron=0, font=(fontStyle, 12), width= 200, compound= 'right', image= DownArrowImg)

    def DeleteDataFunc(x):
        if(x != 'Delete Data'):
            if(x == 'Student'):
                p = DeleteStudent()
                if(p == ''):
                    IncompErrorMsg()
                else:
                    msg.showinfo('Notification', BE.DeleteStudent(p))
            elif(x == 'Marks of a Subject'):
                p = DeleteMarks()
                if(p == ''):
                    IncompErrorMsg()
                else:
                    msg.showinfo('Notification', BE.DeleteSubjectLine(p))
            elif(x == 'Fees of a Year'):
                p = DeleteFees()
                if(p == ''):
                    IncompErrorMsg()
                else:
                    msg.showinfo('Notification', BE.DeleteFeesLine(p))
            clickedDelete.set("Delete Data")

    clickedDelete = StringVar()
    clickedDelete.set('Delete Data')
    DeleteDataList = ['Student', 'Marks of a Subject', 'Fees of a Year']
    DeleteDataOptions = OptionMenu(root, clickedDelete, *DeleteDataList, command= DeleteDataFunc)
    DeleteDataOptions.place(relx = 0.55, rely = 0.5)
    DeleteDataOptions.config(indicatoron=0, font=(fontStyle, 12), width= 200, compound= 'right', image= DownArrowImg)

    def OpenStudentFile():
        level = Toplevel()

        level.title("DST Student")
        level.geometry('400x175+400+400')


        Title = Label(level, text= "Open A Student's File")
        Title.config(font=(fontStyle, 20))
        Title.place(relx=0.3, rely=0.01)


        OpenLabel = Label(level, text= "Enter the Roll no. ")
        OpenLabel.place(relx=0.01, rely=0.25)
        OpenLabel.config(font=(fontStyle, 12))

        RollInput = Entry(level)
        RollInput.place(relx=0.3, rely=0.25, width= 100)
        RollInput.config(font=(fontStyle, 12))
        global roll
        roll = RollInput.get()

        global flag
        flag = True

        def Execute():
            global roll
            roll = RollInput.get()
            if(roll != ''):
                if(roll not in BE.StuList):
                    msg.showinfo('Notification', 'The roll no. does not exist')
                else:
                    os.startfile(DSTFolderPath + 'Students\\' + BE.ConvertRollToText(roll) + '.xlsx')
            level.quit()
            level.destroy()
        
        ExecuteButton = Button(level, text="Execute", font=10, command=Execute)
        ExecuteButton.config(font=(fontStyle, 12))
        ExecuteButton.place(relx=0.8, rely=0.75)

        def Cancel():
            global flag
            flag = False
            level.quit()
            level.destroy()

        CanceleButton = Button(level, text="Cancel", command=Cancel)
        CanceleButton.config(font=(fontStyle, 12))
        CanceleButton.place(relx=0.61, rely=0.75)

        level.mainloop()

    OpenButton = Button(root, text="Open A File", font=2, command=lambda: OpenStudentFile())
    OpenButton.place(relx=0.43, rely=0.75)
    OpenButton.config(font=(fontStyle, 12))
    

    root.mainloop()

main()