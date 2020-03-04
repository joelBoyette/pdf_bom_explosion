
import tkinter as tk
import tkinter.scrolledtext as ScrolledText
import threading
import logging

from log_config import log_location, log_filemode, log_format, log_datefmt

logger = logging.getLogger(__name__)

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


class myGUI(tk.Frame):

    # This class defines the graphical user interface

    def __init__(self, parent, *args, **kwargs):

        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.root = parent
        self.build_gui()

        self.top_level_txtbx = tk.Entry(self.root, width=10)
        self.top_level_txtbx.grid(column=0, row=1, padx=1, pady=1)
        self.btn = tk.Button(self.root, text="Explode BOM", padx=1, pady=1, command=self.more_beef)
        self.btn.grid(column=0, row=5)

    def build_gui(self):

        # Build GUI
        self.root.title('TEST')
        self.root.geometry('1000x200')
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
        logging.basicConfig(filename=log_location + 'dummy.log',
                            filemode=log_filemode,
                            format=log_format,
                            datefmt=log_datefmt,
                            level=logging.CRITICAL)

        # Add the handler to logger
        logger = logging.getLogger()
        logger.addHandler(text_handler)

    def more_beef(self, top_level=None):

        top_level = self.top_level_txtbx.get()

        from dummy_files import dummy_recursion

        t1 = threading.Thread(target=dummy_recursion.beef(int(top_level)))
        t1.start()
        t1.join()


def main():

    root = tk.Tk()
    window = myGUI(root)

    root.mainloop()


if __name__ == '__main__':
    main()

