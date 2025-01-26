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

        self.select_file_button = ctk.CTkButton(
            self.root, text="Выбрать файл", command=self.select_file,
            width=200, height=40, corner_radius=10)
        self.select_file_button.pack(pady=20)

        self.process_data_button = ctk.CTkButton(
            self.root, text="Обработать данные", command=self.process_data,
            width=200, height=40, corner_radius=10)
        self.process_data_button.pack(pady=20)

        self.selected_file = None

    def select_file(self):
        messagebox.showinfo("Информация", "Выберите оригинальный файл счета СМО")
        self.selected_file = filedialog.askopenfilename(
            title="Выберите файл", filetypes=[("Excel files", "*.xlsx;*.xls")])

    def show_save_path_prompt(self):
        messagebox.showinfo("Информация", "Выберите путь для сохранения расшифрованного файла")
        self.choose_save_path()

    def choose_save_path(self):
        folder_path = filedialog.askdirectory(title="Выберите папку для сохранения файла")

        if not folder_path:
            return

        base_filename = "расшифрованный счет"
        filename = f"{base_filename}.xlsx"
        counter = 1

        while os.path.exists(os.path.join(folder_path, filename)):
            filename = f"{base_filename} ({counter}).xlsx"
            counter += 1

        df = pd.read_excel(self.selected_file)

        new_df = df.iloc[:7].copy()
        new_df.insert(4, 'Новый столбец', '')

        for index in range(7, len(df)):
            row_data = df.iloc[index].tolist()[:5]

            if len(row_data) < 5:
                row_data.extend([''] * (5 - len(row_data)))

            key_value = row_data[4]

            if key_value in data_dict:
                row_data.append(data_dict[key_value])
            else:
                row_data.append('')

            if len(df.columns) > 5:
                row_data.append(df.iloc[index, 5])
            else:
                row_data.append('')

            if len(df.columns) > 6:
                row_data.append(df.iloc[index, 6])
            else:
                row_data.append('')

            while len(row_data) < len(new_df.columns):
                row_data.append('')

            new_df.loc[len(new_df)] = row_data

        new_df.to_excel(os.path.join(folder_path, filename), index=False)

        messagebox.showinfo("Успех", "Обработка успешно завершена!")

    def process_data(self):
        if not self.selected_file:
            messagebox.showwarning("Предупреждение", "Сначала выберите файл!")
            return

        self.show_save_path_prompt()


if __name__ == "__main__":
    root = ctk.CTk()
    app = CustomCounterApp(root)
    root.mainloop()
