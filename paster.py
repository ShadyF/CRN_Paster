from tkinter import *
import webbrowser
from pynput.keyboard import Key, Listener, Controller


class App:
    def __init__(self, master):

        self.master = master
        self.CRNs = []
        self.keyboard = None
        self.listener_initialized = False
        master.title("CRN Automatic Paster")

        # Validation
        vcmd = (self.master.register(self.validate), '%S')
        self.v = False  # self.v = IntVar()

        label_greeting = Label(master, text="CRN Automatic Paster")

        self.e1 = Entry(master, width=10, validate="key", validatecommand=vcmd)
        self.e2 = Entry(master, width=10, validate="key", validatecommand=vcmd)
        self.e3 = Entry(master, width=10, validate="key", validatecommand=vcmd)
        self.e4 = Entry(master, width=10, validate="key", validatecommand=vcmd)
        self.e5 = Entry(master, width=10, validate="key", validatecommand=vcmd)
        self.e6 = Entry(master, width=10, validate="key", validatecommand=vcmd)
        self.e7 = Entry(master, width=10, validate="key", validatecommand=vcmd)
        self.e8 = Entry(master, width=10, validate="key", validatecommand=vcmd)
        self.e9 = Entry(master, width=10, validate="key", validatecommand=vcmd)
        self.e10 = Entry(master, width=10, validate="key", validatecommand=vcmd)

        C = Checkbutton(master, text="Press Submit after pasting", command=self.cb)  # variable=self.v
        button_done = Button(master, text="Done", command=self.done_pressed)

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
        button_done.grid(row=3, column=3, columnspan=2, sticky=W + E, pady=5, padx=3)

        self.generate_menu_bar()

    def done_pressed(self):
        self.CRNs = [
            self.e1.get(), self.e2.get(), self.e3.get(), self.e4.get(), self.e5.get(),
            self.e6.get(), self.e7.get(), self.e8.get(), self.e9.get(), self.e10.get()
        ]

        if not self.listener_initialized:
            self.keyboard = Controller()

            listener = Listener(on_release=self.on_release, on_press=None)
            listener.start()

            self.listener_initialized = True

    def generate_menu_bar(self):
        menu = Menu(self.master)
        self.master.config(menu=menu)
        helpmenu = Menu(menu, tearoff=0)
        menu.add_cascade(label="Help", menu=helpmenu)

        helpmenu.add_command(label="How to use", command=self.guide)
        helpmenu.add_command(label="Try it out!", command=self.demo)
        helpmenu.add_command(label="About...", command=self.about)

    def guide(self):
        guide_window = Toplevel()

        v = "1. Copy-Paste or manually input your required CRNs into the program's entry boxes.\n(Keep in mind the " \
            "CRN must not contain any spaces or characters, else they won't be accepted into the entry box)\n2. Press " \
            "the 'Done' Button\n3. Open BSS, highlight/press the FIRST entry box in BSS\n4. Press Shift (Either the " \
            "left or the right one, both work) "

        guide_text = Label(guide_window, text=v, justify=LEFT)
        guide_text.pack()

    def demo(self):
        url = "demo.html"
        webbrowser.open(url, new=2)

    def about(self):
        about_window = Toplevel()

        v = "Made by Shady Fanous\nshady-fanous@aucegypt.edu\nSource code at " \
            "https://github.com/ShadyF/CRN_Paster\nthis tab needs to be redone "

        about_text = Label(about_window, text=v, justify=LEFT)
        about_text.pack()

    def iterate(self):
        for CRN in self.CRNs:
            if CRN:
                self.keyboard.type(CRN)
            self.keyboard.press(Key.tab)
            self.keyboard.release(Key.tab)

        # If Press enter after pasting checkbox is marked
        if self.v:
            self.keyboard.press(Key.enter)
            self.keyboard.release(Key.enter)

    @staticmethod
    def validate(s):
        return s.isdigit()

    def on_release(self, key):
        if key == Key.shift:
            self.iterate()

    def cb(self):
        self.v = not self.v


root = Tk()
app = App(root)
# root.iconbitmap('@clip.ico')
root.mainloop()
