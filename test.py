import tkinter as tk
from tkinter import ttk
from ttkthemes import ThemedTk

class TestApp(ThemedTk):
    def __init__(self):
        super().__init__(theme="arc") # Bắt đầu với theme sáng
        self.title("Kiểm tra các thành phần Tkinter/TTK")
        self.geometry("800x700")

        # Biến để theo dõi trạng thái sáng/tối
        self.dark_mode_var = tk.BooleanVar(value=False) # False = sáng, True = tối

        # Tạo các widget
        self.create_widgets()
        # Áp dụng theme ban đầu
        self.update_theme()

    def create_widgets(self):
        # Frame chính chứa tất cả
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill="both", expand=True)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)

        # --- Cột 1 ---
        col1_frame = ttk.Frame(main_frame)
        col1_frame.grid(row=0, column=0, sticky="nsew", padx=10)

        # LabelFrame cho các widget cơ bản
        basic_widgets_frame = ttk.LabelFrame(col1_frame, text="Các Widget Cơ Bản", padding="10")
        basic_widgets_frame.pack(fill="x", pady=5)

        ttk.Label(basic_widgets_frame, text="Đây là một ttk.Label").pack(pady=5, anchor="w")

        entry = ttk.Entry(basic_widgets_frame)
        entry.insert(0, "Đây là ttk.Entry")
        entry.pack(pady=5, fill="x")

        ttk.Button(basic_widgets_frame, text="Đây là ttk.Button").pack(pady=5)

        ttk.Checkbutton(basic_widgets_frame, text="Đây là ttk.Checkbutton").pack(pady=5, anchor="w")

        # Radiobuttons
        radio_var = tk.StringVar(value="opt1")
        ttk.Radiobutton(basic_widgets_frame, text="Radio Option 1", variable=radio_var, value="opt1").pack(anchor="w")
        ttk.Radiobutton(basic_widgets_frame, text="Radio Option 2", variable=radio_var, value="opt2").pack(anchor="w")

        # Combobox
        combo = ttk.Combobox(basic_widgets_frame, values=["Item A", "Item B", "Item C"], state="readonly")
        combo.set("Đây là ttk.Combobox")
        combo.pack(pady=10, fill="x")

        # --- Cột 2 ---
        col2_frame = ttk.Frame(main_frame)
        col2_frame.grid(row=0, column=1, sticky="nsew", padx=10)

        # Notebook (Tabs)
        notebook = ttk.Notebook(col2_frame)
        notebook.pack(pady=5, fill="both", expand=True)

        tab1 = ttk.Frame(notebook, padding="10")
        tab2 = ttk.Frame(notebook, padding="10")
        notebook.add(tab1, text='Tab 1')
        notebook.add(tab2, text='Tab 2')

        # Nội dung cho Tab 1
        ttk.Label(tab1, text="Nội dung trong Tab 1").pack(pady=5)
        ttk.Scale(tab1, from_=0, to=100, orient="horizontal").pack(pady=10, fill="x")
        ttk.Spinbox(tab1, from_=0, to=10).pack(pady=5)

        # Nội dung cho Tab 2
        ttk.Label(tab2, text="Nội dung trong Tab 2").pack(pady=5)
        ttk.Progressbar(tab2, orient="horizontal", length=200, mode="determinate", value=65).pack(pady=10)

        # Treeview (Bảng)
        tree_frame = ttk.LabelFrame(col1_frame, text="ttk.Treeview", padding="10")
        tree_frame.pack(fill="both", expand=True, pady=5)

        tree = ttk.Treeview(tree_frame, columns=('col1', 'col2'), show='headings')
        tree.heading('col1', text='Cột 1')
        tree.heading('col2', text='Cột 2')
        tree.insert("", "end", values=("Dữ liệu A1", "Dữ liệu A2"))
        tree.insert("", "end", values=("Dữ liệu B1", "Dữ liệu B2"))
        tree.pack(fill="both", expand=True)

        # --- Thanh điều khiển dưới cùng ---
        bottom_frame = ttk.Frame(self, padding="10")
        bottom_frame.pack(fill="x", side="bottom")

        # Nút chuyển đổi giao diện
        self.toggle_button = ttk.Button(bottom_frame, text="Chuyển sang Giao diện Tối", command=self.toggle_dark_mode)
        self.toggle_button.pack()

    def toggle_dark_mode(self):
        """Đảo ngược trạng thái sáng/tối và cập nhật giao diện."""
        self.dark_mode_var.set(not self.dark_mode_var.get())
        self.update_theme()

    def update_theme(self):
        """Cập nhật theme và văn bản của nút dựa trên biến dark_mode_var."""
        if self.dark_mode_var.get():
            # Chuyển sang chế độ tối
            self.set_theme("equilux")
            self.toggle_button.config(text="Chuyển sang Giao diện Sáng")
            # Cập nhật màu nền cho các widget tk cơ bản nếu cần
            self.config(bg="#464646")
            self._update_child_colors(self, bg_color="#464646", fg_color="white")
        else:
            # Chuyển sang chế độ sáng
            self.set_theme("arc")
            self.toggle_button.config(text="Chuyển sang Giao diện Tối")
            # Cập nhật màu nền
            self.config(bg="#f0f0f0")
            self._update_child_colors(self, bg_color="#f0f0f0", fg_color="black")

    def _update_child_colors(self, widget, bg_color, fg_color):
        """Đệ quy cập nhật màu nền cho các widget con."""
        try:
            widget.config(bg=bg_color)
        except tk.TclError:
            pass # Bỏ qua các widget không có thuộc tính 'bg'

        for child in widget.winfo_children():
            self._update_child_colors(child, bg_color, fg_color)

if __name__ == "__main__":
    app = TestApp()
    app.mainloop()

