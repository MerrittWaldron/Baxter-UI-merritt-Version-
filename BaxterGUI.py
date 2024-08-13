"""
Copyright (c) 2022

Permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the "Software"), to deal in
the Software without restriction, including without limitation the rights to
use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies
of the Software, and to permit persons to whom the Software is furnished to do
so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NON INFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
"""
# from tkinter import Tk, Label, Button, messagebox, ttk, HORIZONTAL, Frame, PhotoImage, Entry, LabelFrame,Text, Scrollbar, END

from tkinter import *
from tkinter import simpledialog
from tkinter import ttk, messagebox
import threading
import base64
from processlogs import *
from sharepoint import download_new_files
from sharepoint import upload_new_file
from ekoshandling import *


VERSION = "Baxter Excel Automation V8.2.1"

InputTankLogDir = '/Tank'
InputFilterLogDir = '/Filter'
InputBrewLogDir = '/Brew'
OutputDataFN = '/MergedData'


InputDirectoryPath = './input'
InputOrderDirectoryPath = './inputorder'
InputEkosHopDirectoryPath = './inputhoptracking'
InputTankLogPath = InputDirectoryPath + InputTankLogDir + '/'
InputFilterterLogPath = InputDirectoryPath + InputFilterLogDir + '/'
InputBrewLogPath = InputDirectoryPath + InputBrewLogDir + '/'
OutputDirectoryPath = './output'
OutputDirectoryDataPathFN = OutputDirectoryPath + OutputDataFN + '.xlsx'


check_if_mergeddata_exists(OutputDirectoryDataPathFN)


class Std_redirector(object):
    def __init__(self, widget):
        self.widget = widget

    def write(self, string):
        self.widget.insert(END, string)
        self.widget.see(END)

    def flush(self):
        pass


class MyGUI:
    """ tkinter class for GUI """

    def __init__(self, master):
        self.master = master
        master.title(VERSION)
        self.shareUsername = ''
        self.SharePassword = ''
        self.waitforfinish1 = 0
        self.waitforfinish2 = 0

        self.canvas = Canvas(master, height=900, width=800, bg='#ffa700')
        self.canvas.pack()

        self.frame1 = Frame(master, highlightbackground="black", highlightthickness=4, bd=2, bg='#2c2f33')
        self.frame1.place(relx=0.5, rely=0.02, relwidth=0.95, relheight=0.47, anchor='n')

        #  add png to screen

        self.spi_png_path = './BaxterLogo.png'
        self.master.tk.call('wm', 'iconphoto', self.master._w, PhotoImage(file=self.spi_png_path))

        self.image1 = PhotoImage(file=self.spi_png_path)
        self.label1 = Label(self.frame1, image=self.image1)
        # self.label1.place(relx=0.4, rely=0.01, relwidth=.2, relheight=0.2, anchor='n')
        self.label1.place(relx=0.4, rely=0.01, relwidth=.2, relheight=0.2)

        #  bg='# e6f2ff',
        self.label1A = Label(self.frame1, font=("Calibri", 15), fg='black', bg='orange',
                             text="Click below to make a selection")
        self.label1A.place(relx=.5, rely=0.25, relwidth=.3, relheight=0.1, anchor='n')

        self.optional_button1 = Button(self.frame1, highlightthickness=4, font=("Calibri", 16),
                                       text="Process All Brew Log Files", fg='black', command=self.func1)
        self.optional_button1.place(relx=.5, rely=0.4, relwidth=.4, relheight=0.1, anchor='n')
        self.optional_button1.bind("<Enter>", lambda event: self.optional_button1.configure(fg="orange"))
        self.optional_button1.bind("<Leave>", lambda event: self.optional_button1.configure(fg="black"))

        #  add new Button for Ekos order feature
        self.gen_button1 = Button(self.frame1, highlightthickness=4, font=("Calibri", 16),
                                  text="Order Ingredients from Ekos Report", fg='black',
                                  command=self.func5)
        self.gen_button1.place(relx=.5, rely=0.5, relwidth=.4, relheight=0.1, anchor='n')
        self.gen_button1.bind("<Enter>", lambda event: self.gen_button1.configure(fg="orange"))
        self.gen_button1.bind("<Leave>", lambda event: self.gen_button1.configure(fg="black"))

        #  add new Button for Ekos hop tracking feature
        self.gen_button1 = Button(self.frame1, highlightthickness=4, font=("Calibri", 16),
                                  text="Update Hop tracking Spreadsheet", fg='black',
                                  command=self.func6)
        self.gen_button1.place(relx=.5, rely=0.6, relwidth=.4, relheight=0.1, anchor='n')
        self.gen_button1.bind("<Enter>", lambda event: self.gen_button1.configure(fg="orange"))
        self.gen_button1.bind("<Leave>", lambda event: self.gen_button1.configure(fg="black"))

        self.progress1 = ttk.Progressbar(self.frame1, orient=HORIZONTAL, mode='determinate')
        self.progress1.place(relx=.5, rely=0.82, relwidth=.74, relheight=0.04, anchor='n')
        #
        self.close_button1 = Button(self.frame1, text="Quit", highlightthickness=4, font=("Calibri", 16),
                                    command=master.quit)
        self.close_button1.place(relx=.3, rely=0.9, relwidth=.35, relheight=0.1, anchor='n')
        self.close_button1.bind("<Enter>", lambda event: self.close_button1.configure(fg="red"))
        self.close_button1.bind("<Leave>", lambda event: self.close_button1.configure(fg="black"))
        #
        self.info_button1 = Button(self.frame1, text="Info", highlightthickness=4, font=("Calibri", 16),
                                   command=self.info)
        self.info_button1.place(relx=.7, rely=0.9, relwidth=.35, relheight=0.1, anchor='n')
        self.info_button1.bind("<Enter>", lambda event: self.info_button1.configure(fg="orange"))
        self.info_button1.bind("<Leave>", lambda event: self.info_button1.configure(fg="black"))
        #
        #  create new frame that will contain output text frame with scrollbar
        self.frame2 = Frame(master, highlightbackground="black", highlightcolor="black", highlightthickness=4, bg='#2c2f33')
        self.frame2.place(relx=0.5, rely=0.5, relwidth=.95, relheight=0.48, anchor='n')

        #  bg='# e6f2ff',
        self.label2 = Label(self.frame2, font=("Calibri", 15), fg='black', text="Program Output")
        self.label2.place(relx=0.482, rely=0, relwidth=.95, relheight=0.05, anchor='n')
        #
        #  create a Scrollbar and associate it with txt
        self.scrollb2 = Scrollbar(self.frame2)
        self.scrollb2.pack(side='right', fill='y')

        #  create a Text widget
        self.txt2 = Text(self.frame2, font=("Calibri", 12), borderwidth=3, wrap='word', undo=True,
                         yscrollcommand=self.scrollb2.set)
        self.txt2.place(relx=0.01, rely=0.07, relwidth=.95, relheight=0.9)
        self.scrollb2.config(command=self.txt2.yview)

    def func1(self):
        """ handle step2 functions """
        def step2_thread():
            self.waitforfinish1 = 0
            while self.waitforfinish1 == 1:
                time.sleep(.5)
            self.disable_buttons()
            self.progress1.start(100)
            retok = download_new_files("Brew", self.shareUsername, self.sharePassword)
            if not retok:
                self.progress1.stop()
                self.close_button1['state'] = 'normal'
                messagebox.showerror("Error", "Error Loging into Sharepoint")
                self.enable_buttons()
                self.waitforfinish2 = 0
                return 0
            retok = merge_excel_brew_files(InputBrewLogPath, OutputDirectoryDataPathFN)
            if not retok:
                self.progress1.stop()
                self.close_button1['state'] = 'normal'
                messagebox.showerror("Error Executing Brew Log Scan")
                self.enable_buttons()
                self.waitforfinish2 = 0
                return 0
            self.progress1.stop()
            messagebox.showinfo("", "Brew merged completed successfully")
            self.enable_buttons()
            self.waitforfinish2 = 0
        threading.Thread(target=step2_thread).start()

    def func5(self):
        """ handle processing Ekos ordering functions """
        self.progress1.start(100)
        retok = create_ingredients_order_csv(InputOrderDirectoryPath, OutputDirectoryPath)
        if retok:
            messagebox.showinfo("Information", "Order file Created")
        else:
            messagebox.showinfo("Information", "Did Not Created new order file")
        self.progress1.stop()
        self.enable_buttons()
        self.waitforfinish1 = 0

    """ handle updating hop tracking sheet in SharePoint functions """
    def func6(self):
        self.progress1.start(100)
        retok = download_new_files("Hops", self.shareUsername, self.sharePassword)
        if not retok:
            self.progress1.stop()
            self.close_button1['state'] = 'normal'
            messagebox.showerror("Error", "Error Logging into SharePoint")
            self.enable_buttons()
            return 0
        retok = update_hop_tracking_csv(InputEkosHopDirectoryPath, OutputDirectoryPath)
        if not retok:
            messagebox.showinfo("Information", "Could not open Ekos hop report")
            self.progress1.stop()
            return 0

        sharepointRelativePath = 'Shared Documents/Brewery and Cellar/Brewing Logs/Hops Tracking/'
        sharepointFilename = 'Hops Alpha Worksheet.xlsx'
        newFilePath = './output/'
        #     Parameters:
        #         1st = ./output/testSPhops.xlsx - directory and filename to upload
        #         2nd = 'Shared Documents/Brewery and Cellar/Brewing Logs/Hops Tracking/'
        #               - relative path to where file is going no forward slash at the beginning
        #         3rd = 'testSPhops.xlsx' - filename to upload
        retok = upload_new_file(newFilePath + sharepointFilename, sharepointRelativePath, sharepointFilename, self.shareUsername, self.sharePassword)
        if not retok:
            self.progress1.stop()
            self.close_button1['state'] = 'normal'
            messagebox.showerror("Error", "Error Logging into SharePoint")
            self.enable_buttons()
            return 0
        self.enable_buttons()
        self.progress1.stop()
        messagebox.showinfo("Information", "SharePoint Hop Sheet Updated")

    def info(self):
        print("\nMerge Excel files placed in appropriate input directories \n Resulting output files will appear in output directory")

    def disable_buttons(self):
        """ make sure to disable button while processing """
        # self.gen_button1['state'] = 'disabled'
        # self.create_button1['state'] = 'disabled'
        # self.update_button1['state'] = 'disabled'
        self.optional_button1['state'] = 'disabled'
        self.close_button1['state'] = 'disabled'

    def enable_buttons(self):
        """ make sure to enabled button after processing """
        # self.gen_button1['state'] = 'normal'
        # self.create_button1['state'] = 'normal'
        # self.update_button1['state'] = 'normal'
        self.optional_button1['state'] = 'normal'
        self.close_button1['state'] = 'normal'


def encode(key, clear):
    """ encode information for security """
    enc = []
    for i in range(len(clear)):
        key_c = key[i % len(key)]
        enc_c = chr((ord(clear[i]) + ord(key_c)) % 256)
        enc.append(enc_c)
    return base64.urlsafe_b64encode("".join(enc).encode()).decode()


def decode(key, enc):
    """ decode information for security """

    dec = []
    enc = base64.urlsafe_b64decode(enc).decode()
    for i in range(len(enc)):
        key_c = key[i % len(key)]
        dec_c = chr((256 + ord(enc[i]) - ord(key_c)) % 256)
        dec.append(dec_c)
    return "".join(dec)


def encode_user_credentials(filename, crypt_key, token, username, password):
    """ take clear text and encode """
    credentials = []
    crypt_token = encode(crypt_key, token)
    crypt_username = encode(crypt_key, username)
    crypt_password = encode(crypt_key, password)
    credentials.append(crypt_token)
    credentials.append(crypt_username)
    credentials.append(crypt_password)
    try:
        f = open(filename, 'w')
    except FileNotFoundError:
        print("Cannot open file:" + filename)
        sys.exit(1)
    except Exception as e:
        errno, strerror = e.args
        print(f"I/O error {errno} : {strerror} {filename}")
        sys.exit(1)
    else:
        for line in credentials:
            f.write(line + '\n')
        f.close()


def decode_user_credentials(filename, crypt_key):
    """ take encrypted text and return clear text """

    clear_list = []
    try:
        f = open(filename, 'r')
    except Exception as e:
        errno, strerror = e.args
        print(f"I/O error({errno}): {strerror} {filename}")
        sys.exit(1)
    else:
        for line in f.readlines():
            try:
                clear_list.append(decode(crypt_key, line))
            except Exception as e:
                print("credentials.txt is possibly corrupted. Please delete file 'credentials.txt' and restart program")
                print(e)
        f.close()
        return clear_list


def main():
    """ main function called by script """
    print("Python Version from is " + platform.python_version())
    print("System Version is " + platform.platform())
    print(VERSION)
    localtime = time.asctime(time.localtime(time.time()))
    print("Local current time :", localtime)
    root = Tk()
    root.geometry("1000x600")
    root.resizable(0, 0)  # Don't allow resizing in the x or y direction

    my_gui = MyGUI(root)
    #  redirecting output from script to Tkinter Text window
    sys.stdout = Std_redirector(my_gui.txt2)

    crypt_key = 'secret BAXTER message'
    filename = "./credentials.txt"
    if os.path.isfile(filename):
        user_credentials = decode_user_credentials(filename, crypt_key)
        # print(user_credentials)
        my_gui.shareUsername = user_credentials[1]
        my_gui.sharePassword = user_credentials[2]
    else:
        #  the input dialog
        my_gui.shareUsername = simpledialog.askstring(title="Username",
                                                      prompt="What's your Sharepoint Username?:")
        my_gui.sharePassword = simpledialog.askstring(title="Password",
                                                      prompt="What's your Sharepoint Password?:", show='*')
        token = "ABXVFR"
        encode_user_credentials(filename, crypt_key, token, my_gui.shareUsername, my_gui.sharePassword)

    root.mainloop()
    #  To stop redirecting stdout:
    sys.stdout = sys.__stdout__
    root.destroy()


if __name__ == "__main__":
    main()
