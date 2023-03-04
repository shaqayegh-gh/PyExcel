import pytest
import string
import random
# from pyexcel.xlwt_create import ExcelCreator
from pyexcel.openpyxl_method import OpenPyExcelCreator


class TestUser:

    @pytest.fixture(scope="class")
    def create_json_data_list(self):
        headers = {
            'first_name': 'نام', 'last_name': 'نام خانوادگی', 'age': 'سن', 'phone_number': 'شماره تلفن',
            'address': 'ادرس', 'city_code': 'کد شهر', 'phone': 'تلفن ثابت', 'gender': 'جنسیت'
        }
        data = []
        for i in range(5000):
            first_name = ''.join(random.choices(string.ascii_letters + string.digits, k=10))
            last_name = ''.join(random.choices(string.ascii_letters + string.digits, k=10))
            phone_number = '09' + ''.join(random.choices(string.digits, k=9))
            address = ''.join(random.choices(string.ascii_lowercase, k=30))
            phone = ''.join(random.choices(string.digits, k=8))
            gender = random.choice(['خانم', 'آقا'])
            data.append(
                dict(
                    first_name=first_name, last_name=last_name, age=random.randint(1, 100),
                    phone_number=phone_number, address=address, city_code='021', phone=phone, gender=gender
                )
            )
        return headers, data

    def test_openpy_excel(self, create_json_data_list):
        headers, data = create_json_data_list
        excel_obj = OpenPyExcelCreator()
        workbook = excel_obj.create_workbook_single_sheet(headers_dict=headers, input_data=data)
        workbook.save(f'openpy_test.xlsx')
        response = excel_obj.return_excel_template(workbook=workbook, excel_name='test_template')

        # save the response content to a file on disk
        with open('test_template.xlsx', 'wb') as f:
            f.write(response.content)
