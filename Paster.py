from Tkinter import *
import pyperclip
import pythoncom
import pyHook
import win32com.client
import webbrowser

class App(object):
    def __init__(self, master):
    
        self.master = master
        
        master.title("CRN Automatic Paster")
        
        #Validation
        vcmd = (self.master.register(self.Validate), '%S')
        self.v = False #self.v = IntVar()
        
        label_greeting = Label(master, text="CRN Automatic Paster")
        
        self.e1 = Entry(master, width = 10, validate="key", validatecommand=vcmd)
        self.e2 = Entry(master, width = 10, validate="key", validatecommand=vcmd)
        self.e3 = Entry(master, width = 10, validate="key", validatecommand=vcmd)
        self.e4 = Entry(master, width = 10, validate="key", validatecommand=vcmd)
        self.e5 = Entry(master, width = 10, validate="key", validatecommand=vcmd)
        self.e6 = Entry(master, width = 10, validate="key", validatecommand=vcmd)
        self.e7 = Entry(master, width = 10, validate="key", validatecommand=vcmd)
        self.e8 = Entry(master, width = 10, validate="key", validatecommand=vcmd)
        self.e9 = Entry(master, width = 10, validate="key", validatecommand=vcmd)
        self.e10 = Entry(master, width = 10, validate="key", validatecommand=vcmd)
        
        C = Checkbutton(master, text="Press Submit after pasting", command=self.cb) #variable=self.v
        button_done = Button(master, text="Done", command=self.DonePressed)
        
        label_greeting.grid(row=0, column=0, columnspan=10, pady=10)
        
        self.e1.grid(row=1, column=0, sticky=W, padx=7)
        self.e2.grid(row=1, column=1, sticky=W, padx=7)
        self.e3.grid(row=1, column=2, sticky=W, padx=7)
        self.e4.grid(row=1, column=3, sticky=W, padx=7)
        self.e5.grid(row=1, column=4, sticky=W, padx=7)
        self.e6.grid(row=2, column=0, sticky=W, padx=7, pady=10)
        self.e7.grid(row=2, column=1, sticky=W, padx=7)
        self.e8.grid(row=2, column=2, sticky=W, padx=7)
        self.e9.grid(row=2, column=3, sticky=W, padx=7)
        self.e10.grid(row=2, column=4, sticky=W, padx=7)
        
        C.grid(row=3, column=0, columnspan=2)
        button_done.grid(row=3, column=3, columnspan=2, sticky=W+E, pady=5, padx=3)
        
        self.GenerateMenuBar()
        
        #Setting up pyHook
        self.hm = pyHook.HookManager()
        self.hm.HookKeyboard()
        pythoncom.PumpWaitingMessages()
        
        self.wsh = win32com.client.Dispatch("WScript.Shell")
        
    def DonePressed(self):
        self.iteration = 0
        self.CRNs = [
                     self.e1.get(), self.e2.get(), self.e3.get(), self.e4.get(), self.e5.get(),
                     self.e6.get(), self.e7.get(), self.e8.get(), self.e9.get(), self.e10.get()
                     ]
                     
        #Activates pyHook(Begins detection of keystrokes)
        self.hm.KeyDown = self.OnKeyboardEvent
        
    def GenerateMenuBar(self):
        menu = Menu(self.master)
        self.master.config(menu = menu)
        helpmenu = Menu(menu, tearoff=0)
        menu.add_cascade(label="Help", menu=helpmenu)
        
        helpmenu.add_command(label="How to use", command=self.Guide)
        helpmenu.add_command(label="Try it out!", command=self.Demo)
        helpmenu.add_command(label="About...", command=self.About)
        
    def Guide(self):
        guide_window = Toplevel()
        
        v ="""
        1. Copy-Paste or manually input your required CRNs into the program's entry boxes.
        (Keep in mind the CRN must not contain any spaces or characters, else they won't be accepted into the entry box)
        2. Press the "Done" Button
        3. Open BSS, highlight/press the first box in BSS
        4. Press Shift (Either the left or the right one, both work)
        5. All your CRNs should be pasted into BSS now, press Submit and pray you're not too late.
        """
        
        guide_text = Label(guide_window, text=v, justify=LEFT)
        guide_text.pack()
        
    def Demo(self):
        url = "Demo.html"
        webbrowser.open(url, new=2)
        
    def About(self):
        # about_window = Toplevel()
        # frame = Frame(about_window, width=90, height=90)
        # image = PhotoImage(file="clip.pbm")
        # photo = Label(frame, image=image)
        # photo.image = image #keep a reference!
        # frame.pack()
        # photo.pack()
        pass
        
    def Iterate(self):
        if self.iteration <= (len(self.CRNs) - 1):
            pyperclip.copy(self.CRNs[self.iteration])
            self.wsh.SendKeys(pyperclip.paste())
            self.wsh.SendKeys("{TAB}")
            self.iteration = self.iteration + 1
            self.Iterate()
            
        else:
            #For some reason self.v.get() makes the program crashes when put
            #in this block and the if block, any other block works. Used flag system to bypass
            if self.v:
                self.wsh.SendKeys("{Enter}")
            self.iteration = 0
            
        
    def OnKeyboardEvent(self, event):
        if event.Key == "Lshift" or event.Key == "Rshift":
            self.Iterate()
        #must return true else won't work
        return True
        
    def Validate(self, S):
        return S.isdigit()
        
    def cb(self):
        self.v = not self.v
               
root = Tk()
app = App(root)
#root.iconbitmap('clip.ico')
root.mainloop()