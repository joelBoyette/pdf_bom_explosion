import pandas as pd
import re
import numpy as np
from tkinter import *
# user_top_level = input("Please Enter Top Level")
# user_qty = input("Please enter a qty")
#
# print(f'user top level {user_top_level} and qty {user_qty}')





def main():

    window = Tk()
    window.geometry('350x200')

    top_level_lbl = Label(window, text='Top Level', padx=20, pady=10).grid(column=0, row=0)
    top_level_txtbx = Entry(window, width=10)
    top_level_txtbx.grid(column=0, row=1, padx=1, pady=1)

    qty_lbl = Label(window, text='Make Qty', padx=20, pady=10).grid(column=0, row=2)
    qty_txtbx = Entry(window, width=10)
    qty_txtbx.grid(column=0, row=3, padx=1, pady=1)

    def explode_assembly():
        top_level_input = top_level_txtbx.get()
        qty_input = qty_txtbx.get()
        print(f'part {top_level_input} and qty {qty_input}')

    btn = Button(window, text="Explode BOM", padx=1, pady=1, command=explode_assembly)
    btn.grid(column=0, row=5)

    window.mainloop()


if __name__ == '__main__':
    main()
