import os.path
import sys

import customtkinter as ctk

from tkinter import filedialog, END

import pandas as pd


def open_fils_to_path(name):
    filepaths = filedialog.askopenfilenames(
        title=f"Выберите Excel файлы для {name}",
        filetypes=(("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*"))
    )
    if not filepaths:
        return
    return filepaths

def send_notification(title, message, settime=15):
    plyer.notification.notify(title=title, message=message, app_name="Bam Manager", timeout=settime,
                              app_icon=resource_path(r"static/ico/bam_manager.ico"))

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.normpath(os.path.join(base_path, relative_path))

class AppGui(ctk.CTk):
    def __init__(self):
        self.geomitri_constants()
        super().__init__()
        self.title("Dse Order Manadger")
        self.geometry(f"{self.window_main_x}x{self.window_main_y}")
        ctk.set_appearance_mode("dark")

        self.management_window()

    def geomitri_constants(self):
        self.window_main_x = 700
        self.window_main_y = 420

        """ indent's """
        self.indent_self = 15
        self.indent_frame = 5

        """ main frame """
        self.height_main_frame = 80
        self.height_row_in_frame = 30

        self.width_path_entry = 435
        self.width_name_entry = 149

        self.width_open_button = 50

    def management_window(self):
        main_frame = ctk.CTkFrame(
            self
            , width=self.window_main_x - 2 * self.indent_self
            , height=self.height_main_frame
        )
        main_frame.place(x=self.indent_self, y=self.indent_self)

        self.reply_name_entry = ctk.CTkEntry(
            main_frame
            , width=self.width_name_entry
            , height=self.height_row_in_frame
            , corner_radius=4
            , placeholder_text='Имя файла'
            , state='readonly'
        )
        self.reply_name_entry.place(x=5, y=5)
        self.reply_name_entry.configure(text_color='#9aa5aa', state='normal')
        self.reply_name_entry.delete(0, END)
        self.reply_name_entry.insert(0, 'Файл')
        self.reply_name_entry.configure(state='readonly')

        self.reply_path_entry = ctk.CTkEntry(
            main_frame
            , width=self.width_path_entry
            , height=self.height_row_in_frame
            , corner_radius=4
            , placeholder_text='Введите путь к файлу/файлам отчетов'
        )
        self.reply_path_entry.place(x=self.width_name_entry + 2 * self.indent_frame, y=5)

        self.button_open_folder_reply = ctk.CTkButton(
            main_frame
            , text='Открыть'
            , width=self.width_open_button
            , height=self.height_row_in_frame
            , command=lambda: self.button_path_commands(label_batton='reply')
        )
        self.button_open_folder_reply.place(x=self.width_name_entry + self.width_path_entry + 3 * self.indent_frame,
                                            y=self.indent_frame)

        self.start_button = ctk.CTkButton(
            main_frame
            , width=100
            , height=self.height_row_in_frame
            , text="Начать"
            , fg_color="green"
            , hover_color="darkgreen"
            , command=self.run_manager_thread
        )
        self.start_button.place(x=self.width_name_entry + self.width_path_entry,
                                y=2 * self.height_row_in_frame + 2 * self.indent_frame)

        self.batton_open_result_tabl = ctk.CTkButton(
            main_frame
            ,width=100
            ,height=self.height_row_in_frame
            ,text="Открыть результат"
            ,command=self.command_batton_open_result
            ,fg_color='#b69765'
            ,hover_color='#8f764f'
        )

        logs_frame = ctk.CTkFrame(
            self
            , width=self.window_main_x - 2 * self.indent_self
            , height=self.window_main_y - 3 * self.indent_self - self.height_main_frame
        )
        logs_frame.place(x=self.indent_self, y=2 * self.indent_self + self.height_main_frame)

        self.status_text = ctk.CTkTextbox(
            logs_frame
            , width=self.window_main_x - 2 * self.indent_self - 2 * self.indent_frame
            , height=self.window_main_y - 3 * self.indent_self - self.height_main_frame - 2 * self.indent_frame
        )
        self.status_text.place(x=self.indent_frame, y=self.indent_frame)
        self.status_text.insert("0.0", "Готов к запуску...\n")

    def button_path_commands(self, label_batton: str):
        if label_batton == 'reply':
            # noinspection PyTypeChecker
            path_list_filr = list(open_fils_to_path(name='отчетов'))

            str_paths = ""
            for path in path_list_filr: str_paths += f"{path}, "
            str_paths = str_paths[:-2]

            self.reply_path_entry.delete(0, END)
            self.reply_path_entry.insert(0, str_paths)

            self.reply_name_entry.configure(text_color='#fff', state='normal')
            self.reply_name_entry.delete(0, END)
            self.reply_name_entry.insert(0, os.path.basename(str_paths))
            self.reply_name_entry.configure(state='readonly')

            self.log(f"<Установлен путь для файла отчетов>", color_log='#9aa5aa')

    def log(self, message, color_log=None):
        """Вывод логов в текстовое поле GUI с цветом"""
        self.status_text.insert("end", f"{message}\n")

        if color_log:
            # Получаем позицию только что вставленной строки
            end_index = self.status_text.index("end-1c")
            line_num = end_index.split('.')[0]
            start_pos = f"{int(line_num) - 1}.0"
            end_pos = f"{int(line_num) - 1}.end"

            # Создаём/настраиваем тег и применяем
            tag_name = f"color_{color_log}"
            self.status_text.tag_config(tag_name, foreground=color_log)
            self.status_text.tag_add(tag_name, start_pos, end_pos)

        self.status_text.see("end")

    def run_manager_thread(self):
        """Запуск в отдельном потоке, чтобы GUI не зависал"""
        self.batton_open_result_tabl.place_forget()
        self.start_button.configure(state="disabled")
        self.log("Запуск программы...")
        if pd.isna(self.reply_path_entry.get()):
            self.log("Ошибка, укажите путь к файлу", color_log="red")
            self.start_button.configure(state="normal")
            return

        thread = threading.Thread(target=self.execute_logic, daemon=True)
        thread.start()

    def execute_logic(self):
        self.path_outfile = None
        try:
            # manager = LogicManage_programm()

            if hasattr(manager, 'main'):
                self.path_outfile = manager.main()
            self.batton_open_result_tabl.place(x=self.size_window_x - self.main_frame_pad * 2 - 140 - 5, y=110 - 30 - 5)

            self.log("Complete!", color_log="green")
            self.log("Процесс успешно завершен.", color_log="green")
            send_notification("Программа завершена", "Программа завершена, проверте файл", 16)
            self.start_button.configure(state="normal")
        except Exception as e:
            self.log(f"ERROR: {str(e)}", color_log="red")
        finally:
            self.start_button.configure(state="normal")

if __name__ == "__main__":
    app = AppGui()
    app.mainloop()
