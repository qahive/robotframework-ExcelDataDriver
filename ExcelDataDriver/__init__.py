# Copyright 2020
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#    http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
#
# This file derived from
# Derived from https://github.com/Snooz82/robotframework-datadriver/blob/master/src/DataDriver/DataDriver.py
"""
Enhance robot framework document:
    https://github.com/robotframework/robotframework/blob/master/doc/userguide/src/ExtendingRobotFramework/ListenerInterface.rst
    https://github.com/robotframework/robotframework/tree/master/doc/userguide/src/ExtendingRobotFramework
"""
import glob
import re
import sys
import os.path
import shutil
from copy import deepcopy
from datetime import datetime
from robot.libraries.BuiltIn import BuiltIn
from ExcelDataDriver.Util.CustomSelenium import CustomSelenium
from ExcelDataDriver.ExcelParser import ParserContext
from ExcelDataDriver.ExcelParser.DefaultParserStrategy import DefaultParserStrategy
from ExcelDataDriver.Util.OpenpyxlHelper import OpenpyxlHelper
from ExcelDataDriver.base.robotlibcore import keyword
from ExcelDataDriver.ExcelParser.ExcelTestDataService import ExcelTestDataService
from ExcelDataDriver.ExcelTestDataRow.TestStatus import TEST_STATUSES
from ExcelDataDriver.Keywords.CoreExcelKeywords import CoreExcelKeywords
from ExcelDataDriver.Config.CaptureScreenShotOption import CaptureScreenShotOption


__version__ = '1.0.0'


class ExcelDataDriver:
    """
    ExcelDataDriver is a Robotic Process Automation library (RPA) for RobotFramework that allow the developer create RPA script easier and reduce complexity under robot script layer.

    """
    ROBOT_LISTENER_API_VERSION = 3
    ROBOT_LIBRARY_SCOPE = 'TEST SUITE'
    REPORT_PATH = './/RPA_report'

    # Static properties
    select_test_data = None
    parser_strategy = None

    def __init__(self, file=None, custom_parser=None, capture_screenshot='Always', manually_test=False, validate_data_only=False):
        """ExcelDataDriver can be imported with several optional arguments.

        - ``file``: Excel xlsx test data file.
        - ``custom_parser``: Default will use 'DefaultParserStrategy'.
        - ``capture_screenshot``: Config capture screen shot strategy. Option (Always, OnFailed, Skip) Default (Always).
        - ``validate_data_only``: For only validate the data in the excel file should be valid

        Run ARGS:
        - ``OFFSET_ROW``: Start from test data row. (default is 0)
        - ``MAXIMUM_ROW``: Maximum test data row. (default is None)

        """
        self.ROBOT_LIBRARY_LISTENER = self
        self.file = file

        self.custom_parser = DefaultParserStrategy()
        if custom_parser is not None:
            CustomExcelParser = self.load_module(custom_parser)
            self.custom_parser = CustomExcelParser.CustomExcelParser()
        self.capture_screenshot_option = CaptureScreenShotOption[capture_screenshot]
        self.manually_test = manually_test

        self.excelTestDataService = ExcelTestDataService()
        self.suite_source = None
        self.template_test = None
        self.template_keyword = None
        self.data_table = None
        self.excel_test_data_dict = {}
        self.index = None

        self.wb = None
        self.ws_test_datas = {}
        self.total_datas = []
        self.rerun_failed = True
        self.validate_data_only = validate_data_only

        self.screenshot_running_id = 0

    def load_module(self, module):
        """``Important`` using by local library only."""
        # module_path = "mypackage.%s" % module
        module_path = module
        if module_path in sys.modules.keys():
            return sys.modules[module_path]
        return __import__(module_path, fromlist=[module])

    def start_suite(self, suite, result):
        """``Important`` using by local library only."""
        if self.manually_test:
            return
        """
        Called when a test suite starts.
        Data and result are model objects representing the executed test suite and its execution results, respectively.
        """
        # Delete previous test run result
        try:
            shutil.rmtree(ExcelDataDriver.REPORT_PATH)
        except:
            print('Error while deleting directory')

        # Create report folder
        try:
            os.mkdir(self.REPORT_PATH)
        except OSError:
            print("Creation of the directory %s failed" % self.REPORT_PATH)

        self.suite_source = suite.source
        self._create_data_table()
        self.template_test = suite.tests[0]
        self.template_keyword = self._get_template_keyword(suite)
        temp_test_list = list()
        for test_data in self.data_table:
            self._create_test_from_template(test_data)
            self.excel_test_data_dict[self.test.name] = test_data
            temp_test_list.append(self.test)
        suite.tests = temp_test_list

    def end_suite(self, name, attributes):
        """``Important`` using by local library only."""
        if self.manually_test:
            return

        # Save excel test result data
        if self.file is not None:
            self.excelTestDataService.save_report(ExcelDataDriver.REPORT_PATH+'/summary_report.xlsx')
        # Clear previous data
        self.file = None
        self.suite_source = None
        self.template_test = None
        self.template_keyword = None
        self.data_table = None
        self.excel_test_data_dict = {}
        self.index = None

    def start_test(self, data, result):
        """``Important`` using by local library only."""
        if self.manually_test:
            return
        self.excelTestDataService.select_validation_data(self.excel_test_data_dict[str(data)])

    def end_test(self, data, result):
        """``Important`` using by local library only."""
        if self.manually_test:
            return

        # Capture screenshot
        screen_shot = ''
        if self.validate_data_only is False:
            if self.capture_screenshot_option == CaptureScreenShotOption.Always:
                screen_shot = self._capture_screenshot()
            elif self.capture_screenshot_option == CaptureScreenShotOption.OnFailed and result.passed is False:
                screen_shot = self._capture_screenshot()
        # Update test result
        if result.passed:
            self.excelTestDataService.update_test_result(TEST_STATUSES['pass'], None, screen_shot)
        else:
            self.excelTestDataService.update_test_result(TEST_STATUSES['fail'], result.message, screen_shot)

    def _create_data_table(self):
        """
        This function creates a dictionary which contains all data from data file.
        Keys are header names.
        Values are data of this column as array.
        """
        # Load xlsx test data file
        print('Load excel xlsx test data file')
        self.excelTestDataService.load_test_data(self.file, self.custom_parser)
        offset_row = BuiltIn().get_variable_value('${OFFSET_ROW}', 0)
        maximum_row = BuiltIn().get_variable_value('${MAXIMUM_ROW}', None)
        self.data_table = self.excelTestDataService.get_all_test_data(
            rerun_only_failed=False,
            offset_row=offset_row,
            maximum_row=maximum_row)

    def _get_template_keyword(self, suite):
        self.template_test = suite.tests[0]
        if self.template_test.template:
            for keyword in suite.resource.keywords:
                if self._is_same_keyword(keyword.name, self.template_test.template):
                    return keyword
        raise AttributeError('No "Test Template" keyword found for first test case.')

    #############################################
    #
    # Create test case
    #
    #############################################

    def _create_test_from_template(self, test_data):
        self.test = deepcopy(self.template_test)
        self._replace_test_case_name(test_data)
        self._replace_test_case_keywords(test_data)
        self._add_test_case_tags(test_data)

    def _replace_test_case_name(self, test_data):
        keys = re.findall('\$\{(.*?)\}',self.test.name)
        for key in keys:
            self.test.name = self.test.name.replace('${'+key+'}', test_data.get_test_data_property(key))

    def _replace_test_case_keywords(self, test_data):
        self.test.keywords.clear()
        # Test Setup
        if self.template_test.keywords.setup is not None:
            self.test.keywords.create(name=self.template_test.keywords.setup.name,
                                      type='setup',
                                      args=self.template_test.keywords.setup.args)
        # Test Keyword
        self.test.keywords.create(name=self.template_keyword.name,
                                  args=self._get_template_args(test_data))
        # Test Teardown
        if self.template_test.keywords.teardown is not None:
            self.test.keywords.create(name=self.template_test.keywords.teardown.name,
                                      type='teardown',
                                      args=self.template_test.keywords.teardown.args)

    def _add_test_case_tags(self, test_data):
        # print('Set Test case tags')
        for tag in test_data.get_testcase_tags():
            self.test.tags.add(tag)

    def _get_template_args(self, test_data):
        # print('Set Template args')
        return_args = []
        for arg in self.template_keyword.args:
            arg = arg.replace('${', '').replace('}', '')
            print(arg)
            return_args.append(test_data.get_test_data_property(arg))
        return return_args

    #############################################
    #
    # Utility
    #
    #############################################

    def _get_normalized_keyword(self, keyword):
        return keyword.lower().replace(' ', '').replace('_', '')

    def _is_same_keyword(self, first, second):
        return self._get_normalized_keyword(first) == self._get_normalized_keyword(second)

    def _capture_screenshot(self):
        try:
            # Capture selenium screenshot
            custom_selenium = CustomSelenium()
            today = datetime.now()
            date_and_time = today.strftime("%Y%m%d%H%M%S")
            screenshot_name = date_and_time + '_' + str(self.screenshot_running_id) + '.png'
            self.screenshot_running_id += 1
            try:
                os.makedirs('./' + ExcelDataDriver.REPORT_PATH + '/screenshots/')
            except:
                None
            try:
                custom_selenium.capture_full_page_screenshot('./'+ExcelDataDriver.REPORT_PATH+'/screenshots/'+screenshot_name)
            except Exception as e:
                return str(e)
            return '=HYPERLINK(".//screenshots//' + screenshot_name + '","' + screenshot_name + '")'
        except:
            return None

    #############################################
    #
    # Public Utility Keywords
    #
    #############################################
    @keyword
    def update_test_result(self, status, log_message=None, screenshot=None):
        """
        Manual update test result by call this keyword

        Arguments:
            | *statue*      | Pass/Fail                                  |
            | *log_message* | Any message that want to log into log file |
            | *screenshot*  | Screenshot file path                       |

        """
        self.excelTestDataService.update_test_result(status, log_message, screenshot)

    @keyword
    def update_test_result_if_keyword_fail(self, log_message=None, screenshot=None):
        """
        Keyword for help user to auto log the test status if keyword fail

        Arguments:
            | *log_message* | Any message that want to log into log file |
            | *screenshot*  | Screenshot file path                       |

        """
        keyword_status = BuiltIn().get_variable_value('${KEYWORD_STATUS}', None)
        if keyword_status == 'FAIL':
            self.excelTestDataService.update_test_result(keyword_status, log_message, screenshot)

    @keyword
    def merged_excel_report(self, data_type='DefaultParserStrategy'):
        """
        Merged all test data from report folder into summary_report.xlsx under summary_report_folder

        Arguments:
            | *data_type* | Test data type [ DefaultParserStrategy, CustomExcelParser ] default=DefaultParserStrategy |

        Example:
            | *Keywords*                        |  *Parameters*       |   *Parameters*      | *Parameters*      |
            | Merged Postcomrun Excel Report    |            	      |                     |                   |
            | Merged Postcomrun Excel Report    |  CustomExcelParser  |                     |                   |
        """
        print('Merged test report...')

        # - List all xlsx file under report_folder
        reports = list(glob.iglob(os.path.join(ExcelDataDriver.REPORT_PATH, '*.xlsx')))
        if len(reports) == 0:
            raise IOError('No report xlsx found under "' + ExcelDataDriver.REPORT_PATH + '" folder')

        # - Read all test result
        print('-------------------------------------------------------')
        print('Initial ws test datas')
        summary_wb = OpenpyxlHelper.load_excel_file(reports[0], data_only=False, keep_vba=False)
        reports.pop(0)

        print('Init Parser')
        overall_test_status_is_pass = True
        summary_error_message = ''

        CustomExcelParser = self.load_module(data_type)
        parser_strategy =  CustomExcelParser.CustomExcelParser()
        parser_context = ParserContext(parser_strategy)

        print('Parse wb')
        summary_wb_test_datas = parser_context.parse(summary_wb)
        for test_datas in summary_wb_test_datas.values():
            for test_data in test_datas:
                if test_data.is_fail():
                    overall_test_status_is_pass = False
                    summary_error_message += str(test_data.get_log_message()) + '\r\n'

        for report in reports:
            print('Merged ws test datas : '+report)
            wb = OpenpyxlHelper.load_excel_file(report, data_only=False, keep_vba=False)
            CustomExcelParser = self.load_module(data_type)
            parser_strategy = CustomExcelParser.CustomExcelParser()
            parser_context = ParserContext(parser_strategy)
            wb_test_datas = parser_context.parse(wb)

            for index_ws, ws_test_datas in enumerate(wb_test_datas.values()):
                for index_test_data, test_data in enumerate(ws_test_datas):
                    if test_data.is_not_run() is False:
                        list(list(summary_wb_test_datas.values())[index_ws])[index_test_data].update_result(
                            test_data.get_status(),
                            test_data.get_log_message(),
                            test_data.get_screenshot())
                        if test_data.is_fail():
                            overall_test_status_is_pass = False
                            summary_error_message += str(test_data.get_log_message()) + '\r\n'
            wb.close()

        # Save the result to a new excel files.
        summary_file = os.path.join(ExcelDataDriver.REPORT_PATH, 'summary_report.xlsx')
        summary_wb.save(summary_file)

        # - Return summary test result if have test failed
        if overall_test_status_is_pass is False:
            raise AssertionError(summary_error_message)

    ####################################################
    #
    # Manage Excel Keywords
    #
    ####################################################
    @keyword
    def load_test_data(self, filename, data_type=None):
        """
        Load excel test data

        Arguments:
        |  filename (string)    |   The file name string value that will be used to open the excel file to perform tests upon. |
        |  data_type            |   Test data type [If data type is None will use DefaultParserStrategy]                       |

        Examples:
        | *Keywords*           |  *Parameters*                                      |
        | Open Excel           |  C:\\Python27\\XLSXRobotTest\\XLSXRobotTest.xlsx   |

        """
        if data_type is not None:
            CustomExcelParser = self.load_module(data_type)
            custom_parser = CustomExcelParser.CustomExcelParser()
            self.excelTestDataService.load_test_data(filename, custom_parser)
        else:
            self.excelTestDataService.load_test_data(filename, DefaultParserStrategy())

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
        return self.excelTestDataService.get_all_test_data(rerun_only_failed, offset_row, maximum_row)

    ####################################################
    #
    # Default Test Data Keywords
    #
    ####################################################
    @keyword
    def select_validation_data(self, test_data):
        """
        Select specific test data for validate

        Arguments:
            | test_data | Test data object |
        """
        self.excelTestDataService.select_validation_data(test_data)

    @keyword
    def get_test_data_property(self, property_name):
        """
        Returns test data property based on property name

        Arguments:
            | property_name | Test data property name should be lower case                    |
        """
        return self.excelTestDataService.get_test_data_property(property_name)

    @keyword
    def verify_update_data_property(self, property_name, data_type, allow_none=True, *data_list):
        """
        Verify test data property.

        Arguments:
            | property_name | Test data property name should be lower case                    |
            | data_type     | 'any', 'positive number', 'number', 'yes/no', 'list', 'datetime |
            | allow_none    | Allow property tobe None. (Default is allow / True)             |
            | data_list     | Array of data when data_type is list                            |

        Throw exception if failed
        """
        property_value = self.get_test_data_property(property_name)

        # Verify None
        if allow_none == True and property_value == None:
            return True
        try:
            if data_type == 'positive number':
                if not (property_value >= 0):
                    raise AssertionError('')
            elif data_type == 'number':
                if not (property_value >= 0):
                    raise AssertionError('')
            elif data_type == 'yes/no':
                if not (property_value == 'Yes' or property_value == 'No'):
                    raise AssertionError('')
            elif data_type == 'list':
                if not (property_value in data_list):
                    raise AssertionError('')
            elif data_type == 'datetime':
                if not (type(property_value) == datetime):
                    match_date = re.search('\d{1,2}/\d{1,2}/\d{4}', property_value)
                    if match_date is None:
                        raise AssertionError('')
        except:
            raise AssertionError(str(property_name)+' should be '+str(data_type)+' not '+str(property_value))

    @keyword
    def get_test_result(self):
        """
        Returns current test result (Pass/Fail/None)
        """
        return self.excelTestDataService.get_test_result()

    @keyword
    def get_test_log_message(self):
        """
        Returns current test log message
        """
        return self.excelTestDataService.get_test_log_message()

    @keyword
    def get_test_screen_shot(self):
        """
        Returns current test screen shot
        """
        return self.excelTestDataService.get_test_screen_shot()

    @keyword
    def save_report(self, newfile=None):
        """
        Force save report

        Arguments:
            |  newfile | save report file name. |

        """
        self.excelTestDataService.save_report(newfile)
