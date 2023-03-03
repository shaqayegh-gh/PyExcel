import datetime
from collections import OrderedDict

from xlwt import Workbook, XFStyle, Font, Pattern


def col_width(data):
    """
    Calculate the width of a column based on the maximum length of its cells.
    """
    max_len = 0
    for row in data:
        for cell in row:
            if len(str(cell)) > max_len:
                max_len = len(str(cell))
    return (max_len + 2) * 256  # add some padding and convert to units of 1/256th of a character width


class ExcelCreator:

    @staticmethod
    def _validate_input_data(input_data: list, record_type: str):
        if record_type == 'json':
            record_type_obj = [dict, OrderedDict]
        elif record_type == 'list':
            record_type_obj = [list]
        else:
            raise Exception("record_type must be json or list")
        if isinstance(input_data, list):
            if input_data and not all([type(item) in record_type_obj for item in input_data]):
                raise TypeError(f'All items of input data should be the same type of {record_type}')
            return input_data
        raise TypeError('Input data should be a list')

    @staticmethod
    def convert_json_data(headers_dict: dict, input_data: list):
        """
        :param headers_dict: a dict of headers with english keys and translated values
        :param input_data: list of data records
        :return: convert json data type to list of data with values
        """
        headers_list = list(headers_dict.keys())
        data = [list({key: item[key] for key in headers_list if key in item}.values()) for item in input_data]
        return data

    @staticmethod
    def create_new_workbook(encoding='utf-8') -> Workbook:
        wb = Workbook(encoding=encoding)
        return wb

    @staticmethod
    def add_sheet(workbook: Workbook, sheet_name: str) -> Workbook:
        workbook.add_sheet(sheet_name)
        return workbook

    @staticmethod
    def write_data_to_sheet(workbook: Workbook, sheet_name: str, data: list) -> Workbook:
        """
        :param workbook: object of Workbook
        :param sheet_name: name of sheet which is Written
        :param data: data of sheet
        :return: updated Workbook
        """
        worksheet = workbook.get_sheet(sheet=sheet_name)
        for row_index, row in enumerate(data):
            worksheet.write(row_index, row)  # TODO: NOT  COMPLETED
        # data_str = "\n".join("\t".join(str(cell) for cell in row) for row in data)
        # worksheet = workbook.get_sheet(sheet=sheet_name)
        # worksheet.write(0, 0, data_str)
        # set the width of each column to fit the maximum content
        # for j in range(worksheet.ncols):
        #     worksheet.col(j).width = col_width([worksheet.row(i)[j].value for i in range(worksheet.nrows)])
        return workbook

    def clean_data(self, headers_dict: dict, input_data: list, record_type: str):
        """
        :param headers_dict: a dict of headers with english keys and translated values
        :param input_data: list of data records
        :param record_type: define records type: json or list
        :return:
        """
        validated_data = self._validate_input_data(input_data=input_data, record_type=record_type)
        if record_type == 'json':
            validated_data = self.convert_json_data(headers_dict=headers_dict, input_data=validated_data)
        # add headers to first index of data list
        validated_data.insert(0, list(headers_dict.values()))
        return validated_data

    @classmethod
    def get_default_sheet_style(cls):
        # header style
        header_style = XFStyle()
        header_font = Font()
        header_font.bold = True
        # create a pattern object and set its background color to light blue
        header_pattern = Pattern()
        header_pattern.pattern = Pattern.SOLID_PATTERN
        header_pattern.pattern_fore_colour = 0x16
        header_style.font = header_font
        header_style.pattern = header_pattern

        # sheet style
        sheet_style = XFStyle()

        return header_style, sheet_style

    def create_single_sheet_excel(self, headers_dict: dict, input_data: list, record_type: str = 'json',
                                  sheet_name=None, workbook: Workbook = None, encoding='utf-8'):
        sheet_name = sheet_name or str(datetime.date.today())
        workbook = workbook or self.create_new_workbook(encoding=encoding)
        workbook = self.add_sheet(workbook=workbook, sheet_name=sheet_name)
        data = self.clean_data(headers_dict=headers_dict, input_data=input_data, record_type=record_type)
        workbook = self.write_data_to_sheet(workbook=workbook, sheet_name=sheet_name, data=data)
        header_style, sheet_style = self.get_default_sheet_style()
        sheet = workbook.get_sheet(sheet=sheet_name)
        sheet.row(0).set_style(header_style)
        # for row_idx in range(workbook.get_sheet(sheet=sheet_name).nrows):
        #     sheet.row(row_idx).set_style(sheet_style)

        # response = HttpResponse(content_type='application/ms-excel')
        # response['Content-Disposition'] = f'attachment; filename="test.xlsx"'
        # workbook.save(response)
        return workbook

    # def create_multi_sheet_excel(self):
