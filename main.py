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
            df = pd.read_excel(self.selected_file)
            num_columns = len(df.iloc[10].dropna())

            if num_columns == 11:
                header_rows = 10
                check_col = 5
                insert_col = 6
            elif num_columns == 7:
                header_rows = 7
                check_col = 4
                insert_col = 5
            else:
                raise ValueError("Неподдерживаемое количество столбцов в файле.")

            new_df = df.iloc[:header_rows].copy()
            new_df.insert(insert_col, 'Расшифровка услуги', '')

            for idx in range(header_rows, len(df)):
                row = df.iloc[idx].tolist()
                main_data = row[:check_col + 1]
                while len(main_data) < check_col + 1:
                    main_data.append('')

                code = main_data[check_col]
                description = data_dict.get(code, '')

                extra_data = row[check_col + 1:]
                while len(extra_data) < (len(df.columns) - (check_col + 1)):
                    extra_data.append('')

                full_row = main_data + [description] + extra_data

                if len(full_row) == len(new_df.columns):
                    new_df.loc[len(new_df)] = full_row

            folder_path = filedialog.askdirectory(title="Выберите папку для сохранения")
            if not folder_path:
                return

            base_name = "расшифрованный_счет"
            counter = 1
            file_path = os.path.join(folder_path, f"{base_name}.xlsx")
            while os.path.exists(file_path):
                file_path = os.path.join(folder_path, f"{base_name}({counter}).xlsx")
                counter += 1

            new_df.to_excel(file_path, index=False)
            messagebox.showinfo("Успех", f"Файл успешно сохранен по пути:\n{file_path}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка: {str(e)}")


if __name__ == "__main__":
    root = ctk.CTk()
    app = CustomCounterApp(root)
    root.mainloop()