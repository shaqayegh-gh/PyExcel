from collections import OrderedDict


class ExcelCreator:
    def __init__(self, input_data: list, record_type=dict):
        self.headers_dict = headers_dict
        self.input_data = input_data
        self.record_type = record_type
        self.cleaned_data = []

    def validate_input_data(self):
        if isinstance(self.input_data, list):
            if self.input_data and not all(type(item) == self.record_type for item in self.input_data):
                raise TypeError('All items of input data list should be the same type of record type')
            return self.input_data
        raise TypeError('Input data should be a list')

    def convert_dict_data_to_list(self):
        headers_list = list(self.headers_dict.keys())
        data = [list({key: item[key] for key in headers_list if key in item}.values()) for item in self.input_data]
        data_str = "\n".join("\t".join(str(cell) for cell in row) for row in data)
        print(data)
        print(data_str)
        return data

    @staticmethod
    def insert_data_into_sheet(sheet_name, encoding='utf-8'):
        # wb = xlwt.Workbook(encoding=encoding)
        ws = wb.add_sheet(excel_file_name)

    @staticmethod
    def create_new_workbook(excel_name, encoding='utf-8'):
        wb = xlwt.Workbook(encoding=encoding)


    def create(self, headers_dict: dict, input_data: list):
        pass