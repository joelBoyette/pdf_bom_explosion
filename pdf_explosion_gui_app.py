
import threading
import logging
import tkinter as tk
import tkinter.scrolledtext as ScrolledText
from log_config import log_location, log_filemode, log_format, log_datefmt
import refresh_epicor

# pyinstaller to make a distribution file for users to use program
# https://pyinstaller.readthedocs.io/en/stable/usage.html
# pyinstaller --hidden-import pkg_resources.py2_warn pdf_explosion_gui_app.py
# https://stackoverflow.com/questions/37815371/pyinstaller-failed-to-execute-script-pyi-rth-pkgres-and-missing-packages

# https://www.ghostscript.com/download/gsdnld.html
# user must install Ghostscript 9.50 for Windows (64 bit) on machines for exe to work

logger = logging.getLogger()


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
        # self.root.config(bg="#f79d00")

        self.build_gui()

        # creates labels and text boxes

        # Top Level Number
        self.top_level_lbl = tk.Label(self.root, text='Top Level').grid(column=0, row=0, sticky='w', padx=5)
        self.top_level_txtbx = tk.Entry(self.root, width=20)
        self.top_level_txtbx.grid(column=1, row=0, sticky='w', padx=5)

        # Make Qty
        self.qty_lbl = tk.Label(self.root, text='Make Qty').grid(column=0, row=1, sticky='w', padx=5)
        self.qty_txtbx = tk.Entry(self.root, width=20)
        self.qty_txtbx.grid(column=1, row=1, sticky='w', padx=5)

        # Path to PDF's
        self.path_lbl = tk.Label(self.root, text="Path to PDF's (if not latest)", wraplength=80)\
            .grid(column=0, row=2, sticky='w', padx=5)
        self.path_txtbx = tk.Entry(self.root, width=25)
        self.path_txtbx.grid(column=1, row=2, sticky='w', padx=5)

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
                                     bg='green', fg='White').grid(column=3, row=4, padx=1, pady=5, sticky='w')

        # Refresh data
        self.refresh_btn = tk.Button(self.root, text="Refresh Data", command=refresh_epicor.refresh_data,
                                     padx=1, pady=1, bg='Orange', fg='White').grid(column=1, row=4, padx=1, pady=10)

        self.log_lbl = tk.Label(self.root, text='Log Results').grid(column=0, row=5, padx=5, sticky='w')

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
        logger.addHandler(text_handler)

    def explode_assembly_final(self, top_level=None, qty=None,
                               path=r'\\vimage\latest' + '\\',
                               ignore_epicor=False, email_print=False, email_address=''):

        top_level = self.top_level_txtbx.get()
        qty = self.qty_txtbx.get()
        email_print = self.email_print_value.get()
        email_address = self.address_txtbx.get()

        if not top_level:
            error = 'Must enter Top Level'
            logger.critical(error)
        elif not qty:
            error = 'Must enter Qty'
            logger.critical(error)
        elif email_print == 1 and not email_address:
            error = 'Must enter Email Address'
            logger.critical(error)
        elif email_print == 0 and email_address:
            error = 'Must check Email Prints checkbox if Email Address entered'
            logger.critical(error)
        else:
            error = ''

        if not error:

            import traverse_bom

            if not self.path_txtbx.get() == '':
                path = self.path_txtbx.get()

            ignore_epicor = self.epicor_flag_value.get()
            if ignore_epicor == 1:
                ignore_epicor = True
            else:
                ignore_epicor = False

            # if email checked, mark true
            if email_print == 1:
                email_print = True
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


