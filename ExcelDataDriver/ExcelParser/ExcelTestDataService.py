from datetime import datetime
from ExcelDataDriver.ExcelParser import ParserContext
from ExcelDataDriver.ExcelParser.DefaultParserStrategy import DefaultParserStrategy
from ExcelDataDriver.Util.OpenpyxlHelper import OpenpyxlHelper


class ExcelTestDataService(object):

    # Static properties
    select_test_data = None
    parser_strategy = None

    def __init__(self):
        self.wb = None
        self.ws_test_datas = {}
        self.total_datas = []
        self.rerun_failed = True

    ####################################################
    #
    # Manage Excel Keywords
    #
    ####################################################
    def load_test_data(self, filename, custom_parser):
        """
        Load excel test data

        Arguments:
        |  filename (string)    |   The file name string value that will be used to open the excel file to perform tests upon. |
        |  data_type            |   Test data type [Default=DefaultParserStrategy]                                             |

        Examples:
        | *Keywords*           |  *Parameters*                                      |
        | Open Excel           |  C:\\Python27\\XLSXRobotTest\\XLSXRobotTest.xlsx   |

        """
        print(str(datetime.now()) + ': Load test data '+filename)
        self.fileName = filename
        self.custom_parser = custom_parser

        self.ws_test_datas = {}
        self.total_datas = []
        print(str(datetime.now()) + ': Excel loading....')
        self.wb = OpenpyxlHelper.load_excel_file(self.fileName)
        print(str(datetime.now()) + ': Excel load finished.')
        ExcelTestDataService.parser_strategy = custom_parser
        self.parser_context = ParserContext(ExcelTestDataService.parser_strategy)
        self.ws_test_datas = self.parser_context.parse(self.wb)

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
        self.total_datas = list(
            filter(lambda test_data: test_data.is_pass() is False or rerun_only_failed is False, self.total_datas))
        # Cut the array offset and max
        self.total_datas = self.total_datas[offset_row:(maximum_row + offset_row if maximum_row is not None else None)]
        self.__clear_all_test_result()
        return self.total_datas

    ####################################################
    #
    # Default Test Data Keywords
    #
    ####################################################
    def select_validation_data(self, test_data):
        ExcelTestDataService.select_test_data = test_data

    def get_test_data_property(self, property_name):
        return ExcelTestDataService.select_test_data.get_test_data_property(property_name)

    def update_test_property(self, property_name, property_value):
        ExcelTestDataService.select_test_data.set_test_data_property(property_name, property_value)

    def update_test_result(self, status, log_message=None, screenshot=None):
        ExcelTestDataService.select_test_data.update_result(status, log_message, screenshot)

    def get_test_result(self):
        return ExcelTestDataService.select_test_data.get_status()

    def get_test_log_message(self):
        return ExcelTestDataService.select_test_data.get_log_message()

    def get_test_screen_shot(self):
        return ExcelTestDataService.select_test_data.get_screenshot()

    def save_report(self, newfile=None):
        if newfile is None:
            newfile = self.fileName
        OpenpyxlHelper.save_excel_file(newfile, self.wb)
        self.wb = None
        self.ws_test_datas = {}
        ExcelTestDataService.select_test_data = None

    def __clear_all_test_result(self):
        """
        Clear test result for all rerun test cases
        :return:
        """
        print(str(datetime.now())+': Clear test result...')
        for test_data in self.total_datas:
            test_data.clear_test_result()
        print(str(datetime.now())+': Clear test result completed')
