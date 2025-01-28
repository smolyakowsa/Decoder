import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import os
from data import data_dict


class CustomCounterApp:
    def __init__(self, root):
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")

        self.root = root
        self.root.title("Расшифровка услуг")
        self.root.geometry("500x300")

        self.selected_file = None

        self.select_file_button = ctk.CTkButton(
            self.root, text="Выбрать файл", command=self.select_file,
            width=200, height=40, corner_radius=10)
        self.select_file_button.pack(pady=20)

        self.process_data_button = ctk.CTkButton(
            self.root, text="Обработать данные", command=self.process_data,
            width=200, height=40, corner_radius=10)
        self.process_data_button.pack(pady=20)

    def select_file(self):
        self.selected_file = filedialog.askopenfilename(
            title="Выберите файл", filetypes=[("Excel files", "*.xlsx;*.xls")])
        if self.selected_file:
            messagebox.showinfo("Успех", "Файл успешно выбран!")

    def process_data(self):
        if not self.selected_file:
            messagebox.showwarning("Ошибка", "Сначала выберите файл!")
            return

        try:
            df = pd.read_excel(self.selected_file, header=None)

            code_col_index = self.find_code_column(df)
            if code_col_index is None:
                raise ValueError("Столбец с кодами услуг не найден.")

            new_df = df.copy()
            new_df.insert(code_col_index + 1, "Расшифровка услуги", "")

            for idx in range(len(new_df)):
                if idx == 0:
                    continue

                code = new_df.iloc[idx, code_col_index]
                if pd.isna(code) or code == "":
                    continue

                description = data_dict.get(str(code).strip(), "")
                new_df.iloc[idx, code_col_index + 1] = description

            folder_path = filedialog.askdirectory(title="Выберите папку для сохранения")
            if not folder_path:
                return

            base_name = "расшифрованный_счет"
            counter = 1
            file_path = os.path.join(folder_path, f"{base_name}.xlsx")
            while os.path.exists(file_path):
                file_path = os.path.join(folder_path, f"{base_name}({counter}).xlsx")
                counter += 1

            new_df.to_excel(file_path, index=False, header=False)
            messagebox.showinfo("Успех", f"Файл успешно сохранен по пути:\n{file_path}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка: {str(e)}")

    def find_code_column(self, df):
        for col_index in range(len(df.columns)):
            for cell in df[col_index]:
                if isinstance(cell, str) and "код услуги" in cell.lower():
                    return col_index
        return None


if __name__ == "__main__":
    root = ctk.CTk()
    app = CustomCounterApp(root)
    root.mainloop()