
import threading
import logging
import tkinter as tk
import tkinter.scrolledtext as ScrolledText


from log_config import log_location, log_filemode, log_format, log_datefmt


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
        self.build_gui()

        # creates labels and text boxes
        self.top_level_lbl = tk.Label(self.root, text='Top Level', padx=20, pady=10)
        self.top_level_lbl.grid(column=0, row=0)
        self.top_level_txtbx = tk.Entry(self.root, width=10)
        self.top_level_txtbx.grid(column=0, row=1, padx=1, pady=1)
        self.qty_lbl = tk.Label(self.root, text='Make Qty', padx=20, pady=10)
        self.qty_lbl.grid(column=0, row=2)
        self.qty_txtbx = tk.Entry(self.root, width=10)
        self.qty_txtbx.grid(column=0, row=3, padx=1, pady=10)
        self.btn = tk.Button(self.root, text="Explode BOM", padx=1, pady=1, command=self.explode_assembly_final)
        self.btn.grid(column=0, row=5)
        self.log_lbl = tk.Label(self.root, text='Log')
        self.log_lbl.grid(column=0, row=6)

    def build_gui(self):

        # Build GUI
        self.root.title('TEST')
        self.root.geometry('800x600')
        self.root.option_add('*tearOff', 'FALSE')
        self.grid(column=0, row=0, sticky='ew')
        self.grid_columnconfigure(0, weight=1, uniform='a')
        self.grid_columnconfigure(1, weight=1, uniform='a')
        self.grid_columnconfigure(2, weight=1, uniform='a')
        self.grid_columnconfigure(3, weight=1, uniform='a')

        # Add text widget to display logging info
        st = ScrolledText.ScrolledText(self.root, state='disabled')
        st.configure(font='TkFixedFont')
        st.grid(column=0, row=8, sticky='w', columnspan=4)

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

    def explode_assembly_final(self, top_level=None, qty=None):

        import traverse_bom

        top_level = self.top_level_txtbx.get()
        qty = self.qty_txtbx.get()

        explosion_thread = threading.Thread(target=traverse_bom.explode_assembly(top_level, qty))
        explosion_thread.start()
        explosion_thread.join()


def main():

    root = tk.Tk()
    pdf_app_gui = PDFAppGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()


