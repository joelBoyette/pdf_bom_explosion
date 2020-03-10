
import threading
import logging
import tkinter as tk
import tkinter.scrolledtext as ScrolledText
from log_config import log_location, log_filemode, log_format, log_datefmt
import refresh_epicor

import PIL
from PIL import Image, ImageTk

# pyinstaller to make a distribution file for users to use program
# https://pyinstaller.readthedocs.io/en/stable/usage.html

# https://www.ghostscript.com/download/gsdnld.html
# user must install Ghostscript 9.50 for Windows (64 bit) on machines for exe to work


# https://stackoverflow.com/questions/13318742/python-logging-to-tkinter-text-widget
class TextHandler(logging.Handler):
    # This class allows you to log to a Tkinter Text or ScrolledText widget
    # Adapted from Moshe Kaplan: https://gist.github.com/moshekaplan/c425f861de7bbf28ef06

    def __init__(self, text):

        logging.Handler.__init__(self)  # run the regular Handler __init__
        self.text = text                # Store a reference to the Text it will log to

    def emit(self, record):

        msg = self.format(record)

        def append():
            self.text.configure(state='normal')
            self.text.insert(tk.END, msg + '\n')
            self.text.configure(state='disabled')
            self.text.yview(tk.END)     # Autoscroll to the bottom
        # This is necessary because we can't modify the Text from other threads
        self.text.after(0, append)


class PDFAppGUI(tk.Frame):

    # This class defines the graphical user interface

    def __init__(self, parent, *args, **kwargs):

        # initiates and creates frame
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.root = parent
        # self.root.config(bg="#CCE5FF")

        self.build_gui()

        # creates labels and text boxes

        # Top Level Number
        self.top_level_lbl = tk.Label(self.root, text='Top Level').grid(column=0, row=0, sticky='w')
        self.top_level_txtbx = tk.Entry(self.root, width=20)
        self.top_level_txtbx.grid(column=1, row=0, sticky='w')

        # Make Qty
        self.qty_lbl = tk.Label(self.root, text='Make Qty').grid(column=0, row=1, sticky='w')
        self.qty_txtbx = tk.Entry(self.root, width=20)
        self.qty_txtbx.grid(column=1, row=1, sticky='w')

        # Path to PDF's
        self.path_lbl = tk.Label(self.root, text="Path to PDF's (if not latest)", wraplength=80)\
            .grid(column=0, row=2, sticky='w')
        self.path_txtbx = tk.Entry(self.root, width=25)
        self.path_txtbx.grid(column=1, row=2, sticky='w')

        # Ignore Epicor Structure
        self.epicor_flags_lbl = tk.Label(self.root, text="Ignore Epicor Structure").grid(column=3, row=0, sticky='w')
        self.epicor_flag_value = tk.IntVar()
        self.epicor_flags_cb = tk.Checkbutton(self.root, text="", variable=self.epicor_flag_value)
        self.epicor_flags_cb.grid(column=4, row=0, sticky='w')

        # Email Prints
        self.email_print_lbl = tk.Label(self.root, text="Email Prints").grid(column=3, row=1, sticky='w')
        self.email_print_value = tk.IntVar()
        self.email_print_cb = tk.Checkbutton(self.root, text="", variable=self.email_print_value)
        self.email_print_cb.grid(column=4, row=1, sticky='w')

        # Email Address
        self.address_lbl = tk.Label(self.root, text="Email Address").grid(column=3, row=2, sticky='w')
        self.address_txtbx = tk.Entry(self.root, width=30)
        self.address_txtbx.grid(column=4, row=2)

        # Explode Button
        self.explode_btn = tk.Button(self.root, text="Explode BOM", command=self.explode_assembly_final,
                                     padx=1, pady=1, bg='green', fg='White').grid(column=3, row=4, padx=1, pady=10)

        # Refresh data
        self.refresh_btn = tk.Button(self.root, text="Refresh Data", command=refresh_epicor.refresh_data,
                                     padx=1, pady=1, bg='Orange', fg='White').grid(column=1, row=4, padx=1, pady=10)

        # # PDF BOM image
        # # use PIL to convert to format acceptable for tkinter
        # image = PIL.Image.open('pdf_bom.PNG')
        # converted_image = PIL.ImageTk.PhotoImage(image)
        #
        # # add converted image to label. must keep a reference, bug in tkinter
        # # http://effbot.org/pyfaq/why-do-my-tkinter-images-not-appear.htm
        # img_lbl = tk.Label(self.root, image=converted_image)
        # img_lbl.image = converted_image
        # img_lbl.grid(column=2, row=9, columnspan=2, rowspan=2, padx=5, pady=5)

    def build_gui(self):

        # Build GUI
        self.root.title('PDF BOM Explosion/Shortage Report')
        self.root.geometry('700x400')
        self.root.option_add('*tearOff', 'FALSE')
        self.grid(column=0, row=0, sticky='ew')
        self.grid_columnconfigure(0, weight=1, uniform='a')
        self.grid_columnconfigure(1, weight=1, uniform='a')
        self.grid_columnconfigure(2, weight=1, uniform='a')
        self.grid_columnconfigure(3, weight=1, uniform='a')

        # Add text widget to display logging info
        st = ScrolledText.ScrolledText(self.root, state='disabled', width=80, height=10)
        st.grid(column=0, row=8, sticky='w', columnspan=8, padx=5)
        st.configure(font='TkFixedFont')

        # Create textLogger
        text_handler = TextHandler(st)

        # Logging configuration
        logging.basicConfig(filename=log_location + 'app.log',
                            filemode=log_filemode,
                            format=log_format,
                            datefmt=log_datefmt,
                            level=logging.CRITICAL)

        # Add the handler to logger
        logger = logging.getLogger()
        logger.addHandler(text_handler)

    def explode_assembly_final(self, top_level=None, qty=None,
                               path=r'\\vimage\latest' + '\\',
                               ignore_epicor=False, email_print=False, email_address=''):

        import traverse_bom

        top_level = self.top_level_txtbx.get()
        qty = self.qty_txtbx.get()

        if not self.path_txtbx.get() == '':
            path = self.path_txtbx.get()

        ignore_epicor = self.epicor_flag_value.get()
        if ignore_epicor == 1:
            ignore_epicor = True
        else:
            ignore_epicor = False

        # if email supplier check, mark true and get email address
        email_print = self.email_print_value.get()
        if email_print == 1:
            email_print = True
            email_address = self.address_txtbx.get()
        else:
            email_print = False
            email_address = ''

        explosion_thread = threading.Thread(target=traverse_bom.explode_assembly(top_level,
                                                                                 qty,
                                                                                 path,
                                                                                 ignore_epicor,
                                                                                 email_print,
                                                                                 email_address))
        explosion_thread.start()
        explosion_thread.join()


def main():

    root = tk.Tk()

    pdf_app_gui = PDFAppGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()


