import tkinter as tk

root = tk.Tk()
root.title("Dynamic Radiobutton Font Example")

selected_option = tk.StringVar()
selected_option.set("Option 1")

def change_font_size():
    if selected_option.get() == "Option 1":
        radio1.config(font=("Courier", 14))
        radio2.config(font=("Courier", 10))
    else:
        radio1.config(font=("Courier", 10))
        radio2.config(font=("Courier", 14))

radio1 = tk.Radiobutton(
    root,
    text="Tốc độ đã ghi 1",
    variable=selected_option,
    value="Option 1",
    command=change_font_size
)
radio1.pack(pady=5)

radio2 = tk.Radiobutton(
    root,
    text="Tốc độ đã ghi 2",
    variable=selected_option,
    value="Option 2",
    command=change_font_size
)
radio2.pack(pady=5)

# Initial font configuration
radio1.config(font=("Courier", 14))
radio2.config(font=("Courier", 10))

root.mainloop()