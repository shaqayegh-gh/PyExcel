import pytest
import string
import random
from pyexcel.xlwt_create import ExcelCreator


class TestUser:

    @pytest.fixture(scope="class")
    def create_json_data_list(self):
        headers = {
            'first_name': 'نام', 'last_name': 'نام خانوادگی', 'age': 'سن', 'phone_number': 'شماره تلفن',
            'address': 'ادرس'
        }
        data = []
        for i in range(100):
            first_name = ''.join(random.choices(string.ascii_letters + string.digits, k=10))
            last_name = ''.join(random.choices(string.ascii_letters + string.digits, k=10))
            phone_number = '09' + ''.join(random.choices(string.digits, k=9))
            address = ''.join(random.choices(string.ascii_lowercase, k=30))
            data.append(
                dict(
                    first_name=first_name, last_name=last_name, age=random.randint(1, 100),
                    phone_number=phone_number, address=address
                )
            )
        return headers, data

    def test_excel_creation(self, create_json_data_list):
        headers, data = create_json_data_list
        excel_obj = ExcelCreator()
        workbook = excel_obj.create_single_sheet_excel(headers_dict=headers, input_data=data,
                                                       record_type='json')
        workbook.save(f'test.xlsx')
