import tkinter as tk

root = tk.Tk()
root.title("Tkinter Test")
root.geometry("400x300")

label = tk.Label(root, text="Tkinter is working!")
label.pack(pady=50)

root.mainloop()
