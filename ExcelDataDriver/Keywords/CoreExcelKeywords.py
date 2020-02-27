import sys
import importlib
from ExcelDataDriver.ExcelParser.ParserContext import ParserContext
from ExcelDataDriver.ExcelParser.DefaultParserStrategy import DefaultParserStrategy
from ExcelDataDriver.Util.OpenpyxlHelper import OpenpyxlHelper
from ExcelDataDriver.base.robotlibcore import keyword

class CoreExcelKeywords(object):

    ROBOT_LIBRARY_SCOPE = 'GLOBAL'

    # Static properties
    select_test_data = None
    parser_strategy = None

    def __init__(self):
        self.wb = None
        self.ws_test_datas = {}
        self.total_datas = []
        self.rerun_failed = True

        # Reference data
        self.reference_data = dict()

    ####################################################
    #
    # Load reference excel datas
    #
    ####################################################
    @keyword
    def load_reference_data(self, alias_name, filename, custom_parser_module='ExcelDataDriver.ExcelParser.DefaultReferenceParserStrategy', custom_parser_class='DefaultReferenceParserStrategy'):
        """
        Load reference data with specific Parser

        Arguments:
        |  alias_name           |   alias_name for refer to the reference data |
        |  filename (string)    |   The file name string value that will be used to open the excel file to perform tests upon. |
        |  custom_parser_module |   Test data parser module is ExcelDataDriver.ExcelParser.DefaultReferenceParserStrategy |
        |  custom_parser_class  |   Test data parser class is DefaultReferenceParserStrategy |
        """
        reference_wb = OpenpyxlHelper.load_excel_file(filename)
        CustomExcelParser = self._load_customer_class_from_module(custom_parser_module, custom_parser_class)
        parser_context = ParserContext(CustomExcelParser())
        self.reference_data[alias_name] = {
            'selected': None,
            'data': parser_context.parse(reference_wb)
        }
        reference_wb.close()

    def _load_customer_class_from_module(self, module_name, class_name):
        MyClass = getattr(importlib.import_module(module_name), class_name)
        return MyClass

    ####################################################
    #
    # Manage Excel Keywords
    #
    ####################################################
    @keyword
    def load_test_data(self, filename, data_type='DefaultParserStrategy'):
        """
        Load excel test data

        Arguments:
        |  filename (string)    |   The file name string value that will be used to open the excel file to perform tests upon. |
        |  data_type            |   Test data type [Default=DefaultParserStrategy]                    |

        Examples:
        | *Keywords*           |  *Parameters*                                      |
        | Load test data       |  C:\\Python\\XLSXRobotTest\\XLSXRobotTest.xlsx     |

        """
        self.fileName = filename
        self.data_type = data_type

        self.ws_test_datas = {}
        self.total_datas = []
        # self.wb = OpenpyxlHelper.load_excel_file(self.fileName)

        # CoreExcelKeywords.parser_strategy = self.__get_parser_based_on_data_type(data_type)
        # self.parser_context = ParserContext(CoreExcelKeywords.parser_strategy)
        # self.ws_test_datas = self.parser_context.parse(self.wb)

    @keyword
    def get_all_test_data(self, rerun_only_failed=False, offset_row=0, maximum_row=None):
        """
        Get all test datas from current excel.

        Arguments:
            |  rerun_only_failed    |   Rerun only failed case default is False              |
            |  offset_row           |   Number of offset row. default is 0                   |
            |  maximum_row          |   Maximum row record. default is None mean no limit    |

        Examples:
            | *Keywords*           |  *Parameters*      |   *Parameters*      | *Parameters*      |
            | Get all test datas   |                    |                     |                   |
            | Get all test datas   |  ${True}           | ${10}               | ${10}             |
        """
        # Get all test data from worksheet
        if len(self.total_datas) > 0:
            return self.total_datas
        if offset_row is not None:
            offset_row = int(offset_row)
        if maximum_row is not None:
            maximum_row = int(maximum_row)
        for test_datas in self.ws_test_datas.values():
            self.total_datas = self.total_datas + test_datas
        # Filter if rerun_failed case only
        self.total_datas = list(filter(lambda test_data: test_data.is_pass() is False or rerun_only_failed is False, self.total_datas))
        # Cut the array offset and max
        self.total_datas = self.total_datas[offset_row:(maximum_row + offset_row if maximum_row is not None else None)]
        self.__clear_all_test_result()
        return self.total_datas

    def __get_parser_based_on_data_type(self, data_type):
        return DefaultParserStrategy()

    ####################################################
    #
    # Default Test Data Keywords
    #
    ####################################################
    @keyword
    def select_validation_data(self, test_data):
        CoreExcelKeywords.select_test_data = test_data

    @keyword
    def get_test_data_property(self, property_name):
        return CoreExcelKeywords.select_test_data.get_test_data_property(property_name)

    @keyword
    def update_test_result(self, status, screenshot, log_message=""):
        CoreExcelKeywords.select_test_data.update_result(status, log_message, screenshot)

    @keyword
    def get_test_result(self):
        return CoreExcelKeywords.select_test_data.get_status()

    @keyword
    def get_test_log_message(self):
        return CoreExcelKeywords.select_test_data.get_log_message()

    @keyword
    def get_test_screen_shot(self):
        return CoreExcelKeywords.select_test_data.get_screenshot()

    @keyword
    def save_report(self, newfile=None):
        if newfile is None:
            newfile = self.fileName
        OpenpyxlHelper.save_excel_file(newfile, self.wb)
        self.wb = None
        self.ws_test_datas = {}
        CoreExcelKeywords.select_test_data = None

    def __clear_all_test_result(self):
        """
        Clear test result for all rerun test cases
        :return:
        """
        print('Clear test result...')
        for test_data in self.total_datas:
            test_data.clear_test_result()
