import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import os
from data import data_dict  # Импортируем словарь с расшифровками


class CustomCounterApp:
    def __init__(self, root):
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")

        self.root = root
        self.root.title("Расшифровка услуг")
        self.root.geometry("500x300")

        self.selected_file = None

        # Кнопки интерфейса
        self.select_file_button = ctk.CTkButton(
            self.root, text="Выбрать файл", command=self.select_file,
            width=200, height=40, corner_radius=10)
        self.select_file_button.pack(pady=20)

        self.process_data_button = ctk.CTkButton(
            self.root, text="Обработать данные", command=self.process_data,
            width=200, height=40, corner_radius=10)
        self.process_data_button.pack(pady=20)

    def select_file(self):
        """Выбор файла для обработки."""
        self.selected_file = filedialog.askopenfilename(
            title="Выберите файл", filetypes=[("Excel files", "*.xlsx;*.xls")])
        if self.selected_file:
            messagebox.showinfo("Успех", "Файл успешно выбран!")

    def process_data(self):
        """Обработка данных."""
        if not self.selected_file:
            messagebox.showwarning("Ошибка", "Сначала выберите файл!")
            return

        try:
            # Чтение файла
            df = pd.read_excel(self.selected_file)

            # Проверка количества столбцов в 11 строке (индекс 10)
            num_columns = len(df.iloc[10].dropna())  # Убираем пустые значения
            print(f"Количество столбцов в 11 строке: {num_columns}")

            # Определение параметров обработки
            if num_columns == 11:
                # Ресо-Мед
                header_rows = 10  # Шапка до 10 строк (включительно)
                check_col = 5     # Код услуги в 6 столбце (индекс 5)
                insert_col = 6    # Вставка нового столбца после 6 столбца (индекс 6)
                smo_type = "Ресо-Мед"
            elif num_columns == 7:
                # Капитал МС
                header_rows = 7   # Шапка до 7 строк (включительно)
                check_col = 4     # Код услуги в 5 столбце (индекс 4)
                insert_col = 5    # Вставка нового столбца после 5 столбца (индекс 5)
                smo_type = "Капитал МС"
            else:
                raise ValueError(
                    f"Неподдерживаемое количество столбцов в 11 строке: {num_columns}. "
                    "Ожидается 7 (Капитал МС) или 11 (Ресо-Мед)."
                )

            print(f"Тип файла определен как: {smo_type}")

            # Создание нового DataFrame с шапкой
            new_df = df.iloc[:header_rows].copy()
            new_df.insert(insert_col, 'Расшифровка услуги', '')

            # Обработка данных
            for idx in range(header_rows, len(df)):
                row = df.iloc[idx].tolist()

                # Основные данные
                main_data = row[:check_col + 1]
                while len(main_data) < check_col + 1:
                    main_data.append('')

                # Расшифровка услуги
                code = main_data[check_col]
                description = data_dict.get(code, '')

                # Дополнительные колонки
                extra_data = row[check_col + 1:]
                while len(extra_data) < (len(df.columns) - (check_col + 1)):
                    extra_data.append('')

                # Формируем итоговую строку
                full_row = main_data + [description] + extra_data

                # Проверяем, что количество столбцов совпадает
                if len(full_row) == len(new_df.columns):
                    new_df.loc[len(new_df)] = full_row
                else:
                    print(f"Пропущена строка {idx + 1}: несоответствие столбцов")

            # Сохранение файла
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