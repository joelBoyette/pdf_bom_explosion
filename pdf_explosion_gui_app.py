
import threading
import logging
import tkinter as tk
import tkinter.scrolledtext as ScrolledText
from log_config import log_location, log_filemode, log_format, log_datefmt

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
        self.top_level_lbl = tk.Label(self.root, text='Top Level', padx=20, pady=10)
        self.top_level_lbl.grid(column=0, row=0)

        self.top_level_txtbx = tk.Entry(self.root, width=10)
        self.top_level_txtbx.grid(column=0, row=1, padx=1, pady=1)

        self.qty_lbl = tk.Label(self.root, text='Make Qty', padx=20, pady=10)
        self.qty_lbl.grid(column=1, row=0)

        self.qty_txtbx = tk.Entry(self.root, width=10)
        self.qty_txtbx.grid(column=1, row=1, padx=1, pady=10)

        self.path_lbl = tk.Label(self.root, text="Path to PDF's (if not latest)", padx=20, pady=10)
        self.path_lbl.grid(column=2, row=0, padx=1, pady=10)

        self.path_txtbx = tk.Entry(self.root, width=40)
        self.path_txtbx.grid(column=2, row=1, padx=1, pady=10)

        self.epicor_flags_lbl = tk.Label(self.root, text="Ignore Epicor Structure", padx=20, pady=10)
        self.epicor_flags_lbl.grid(column=3, row=0, padx=1, pady=10)

        self.epicor_flag_value = tk.IntVar()
        self.epicor_flags_cb = tk.Checkbutton(self.root, text="", variable=self.epicor_flag_value, padx=90, pady=1)
        self.epicor_flags_cb.grid(column=3, row=1, sticky='w')

        self.btn = tk.Button(self.root, text="Explode BOM", command=self.explode_assembly_final,
                             padx=1, pady=1, bg='green', fg='White')
        self.btn.grid(column=3, row=4, padx=1, pady=10)

    def build_gui(self):

        # Build GUI
        self.root.title('PDF BOM Explosion/Shortage Report')
        self.root.geometry('800x400')
        self.root.option_add('*tearOff', 'FALSE')
        self.grid(column=0, row=0, sticky='ew')
        self.grid_columnconfigure(0, weight=1, uniform='a')
        self.grid_columnconfigure(1, weight=1, uniform='a')
        self.grid_columnconfigure(2, weight=1, uniform='a')
        self.grid_columnconfigure(3, weight=1, uniform='a')

        # Add text widget to display logging info
        st = ScrolledText.ScrolledText(self.root, state='disabled', width=80, height=10)
        st.grid(column=0, row=8, sticky='w', columnspan=8, padx=15)
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

    def explode_assembly_final(self, top_level=None, qty=None, path=r'\\vimage\latest' + '\\', ignore_epicor=False):

        import traverse_bom

        top_level = self.top_level_txtbx.get()
        qty = self.qty_txtbx.get()

        if not self.path_txtbx.get() == '':
            path = self.path_txtbx.get()

        ignore_epicor = self.epicor_flag_value.get()
        if ignore_epicor == 1:
            ignore_epicor = True

        explosion_thread = threading.Thread(target=traverse_bom.explode_assembly(top_level, qty, path, ignore_epicor))
        explosion_thread.start()
        explosion_thread.join()


def main():

    root = tk.Tk()

    pdf_app_gui = PDFAppGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()


