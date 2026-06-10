import pandas as pd

from scripts.handlings.handler_reply import HandlerReplyTabel


class ScriptReplyTabel:
    def __init__(self):
        self.data = {}
        self.receive_data = {}
        self.data_handler = {}
        self.path_file = None

    def main(self, path_reply_tabl_file):
        self.path_file = path_reply_tabl_file

        handler_reply_file = HandlerReplyTabel()
        self.data_handler = handler_reply_file.main(path_to_reply_tabel=path_reply_tabl_file).copy()

        self.filter_for_data()

        return self.data

    def filter_for_data(self):
        list_of_dse_and_name_from_handler_file = []

        abbreviation = ''
        product_abbreviation = ''
        num_row_abb= 0
        for num_row, row_data_handler in self.data_handler.items():
            dse_and_name_halding_file = row_data_handler.get('Номенклатура', '')
            if pd.isna(dse_and_name_halding_file):
                if num_row_abb == num_row - 1:

                    abbreviation = product_abbreviation
                    product_abbreviation = f"{row_data_handler.get('Бух заказ', '')}"
                else:
                    product_abbreviation = f"{row_data_handler.get('Бух заказ', '')}"
                num_row_abb = num_row
                continue
            list_of_dse_and_name_from_handler_file.append(dse_and_name_halding_file)
            dse_name = dse_and_name_halding_file

            if " " in dse_name:
                dse = dse_name[:dse_name.index(' ')]
                name = dse_name[len(dse):]

                self.data[dse_name] = {dse_name:{
                    'Бух заказ': abbreviation
                    ,'Заказ Бух': product_abbreviation
                    ,'Дсе': dse
                    , 'Наименование': name
                    , '': ''
                    , 'Номенклатура': dse_name
                }}



if __name__ == '__main__':
    app = ScriptReplyTabel()
    datas = app.main(r"C:\Users\yakovlev_nd\Desktop\Tests\gfgdgssd\26,06,01-08,31.xlsx")

    from scripts.excel_enter import ExcelDataInserter

    inserter = ExcelDataInserter(r"C:\Users\yakovlev_nd\Desktop\Tests\gfgdgssd\26,06,01-08,31.xlsx")
    inserter.insert_data(datas, sheet_name="Изделия")
