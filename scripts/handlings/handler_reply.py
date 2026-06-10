import sys
import os

sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from scripts.excel_reader import ExcelReader


class HandlerReplyTabel:
    def __init__(self):
        self.data = {}
        self.receive_data = {}

    def main(self, path_to_reply_tabel: str):
        """
        return: Обычн
        """
        self.read_excel= ExcelReader(path_to_reply_tabel, header_row=13)
        self.filter_data()


        return self.data

    def filter_data(self):
        self.data = self.read_excel.get_dict_all_data()




if __name__ == '__main__':
    app = HandlerReplyTabel()
    data = app.main(path_to_reply_tabel=r"C:\Users\yakovlev_nd\Desktop\Tests\DseParser\new\26,01,01-12,31.xlsx")

    for num_,row_ in data.items():
        print(num_,row_)
