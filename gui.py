from tkinter import *

main_window = Tk()
main_window.title("Program do rozliczeń")
main_window.geometry("500x500")


def my_click():
    label=Label(main_window, text="wynik to " + c)
    label.pack()

my_label = Label(main_window, text="hello")
my_button = Button(main_window, text="click me", command=my_click)
e1= Entry(main_window, text="a")

e1.pack()

my_label.pack()
my_button.pack()
main_window.mainloop()