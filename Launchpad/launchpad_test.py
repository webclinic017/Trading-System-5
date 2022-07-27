##same as launchpad, but obkect oriented approach
from launchpad_app import Application
import tkinter as tk

# start the program
root = tk.Tk()
app = Application(master=root)
root.geometry("450x450")
root.title("LAUNCHPAD")
app.mainloop()
