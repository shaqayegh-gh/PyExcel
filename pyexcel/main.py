from collections import OrderedDict

from django.http import HttpResponse
from rest_framework.renderers import JSONRenderer
from django.http import StreamingHttpResponse
import json

class BaseExcelCreator:
    @staticmethod
    def _validate_input_data(input_data: list, record_type: str):
        """
        Check if all items of input data list are the same type
        :param input_data: list of data records
        :param record_type: record type can be dict, orderdict or list
        :return: it will raise error if records are not the same else return data
        """
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
    def return_excel_template(cls, workbook, excel_name: str):
        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = f'attachment; filename={excel_name}.xlsx'
        workbook.save(response)
        return response


class StreamQueryset:
    """
    using streaming for serialzing large querysets
    """
    def __init__(self, queryset, serializer_class):
        self.queryset = queryset
        self.serializer_class = serializer_class

    def stream_queryset(self):
        for obj in self.queryset.iterator():
            serializer = self.serializer_class(obj)
            yield JSONRenderer().render(serializer.data)

    def __call__(self):
        """
        convert json stream to list of jsons
        """
        json_stream = self.stream_queryset()
        json_list = [json.loads(line) for line in json_stream]
        return json_list

