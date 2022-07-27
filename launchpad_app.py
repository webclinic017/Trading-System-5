import tkinter as tk
from tkinter import *
from PIL import Image
from PIL import ImageTk

class Application(tk.Frame):
    
    def __init__(self, master=None):
        super().__init__(height=450, width=450)
        self.master = master
        self.pack()
        self.createWidgets()

    def createWidgets(self):
        self.quit = tk.Button(self, text="Exit", fg="red", command=self.master.destroy)
        self.quit.place(x=300, y=340)
        #ticker entry/buy sell
        self.side = IntVar()
        self.ticker = StringVar()
        self.stop = IntVar()
        self.long = Radiobutton(self, text="Buy", variable=self.side, value=1, command=self.buttonCallback)
        self.long.place(x=300, y=290)
        self.short = Radiobutton(self, text="Sell", variable=self.side, value=2, command=self.buttonCallback)
        self.short.place(x=300, y=310)
        self.entry = Entry(self, textvariable=self.ticker, width=10)
        self.entry.place(x=150, y=290)
        self.stop = Entry(self, textvariable=self.stop, width=10)
        self.stop.place(x=150, y=350)
        ticker_label = Label(self, text="Enter a Ticker:")
        ticker_label.place(x=150, y=270)
        stop_label = Label(self, text="Enter Stop Price:")
        stop_label.place(x=150, y=330)
        #image
        width, height = 250, 250
        self.img = Image.open("launchpad.gif")
        self.img = self.img.resize((width, height), Image.ANTIALIAS)
        render = ImageTk.PhotoImage(self.img)
        self.img = Label(image=render)
        self.img.image = render
        self.img.place(x=100, y=0)

    def buttonCallback(self):
        self.radio_selection = self.side.get()
        ticker = self.entry.get().upper()
        stop = self.stop.get()
        print("ticker selected: ", ticker)
        print("stop price: ", stop)
        if (self.radio_selection == 1):
            print("buy order")
        else:
            print("sell order")
        return


        
