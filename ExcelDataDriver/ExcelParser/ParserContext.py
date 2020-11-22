from datetime import datetime
from collections import OrderedDict


class ParserContext:

    def __init__(self, parser_strategy):
        self.parser_strategy = parser_strategy

    def parse(self, wb):
        """
        Parse excel test data row based on Excel Parser Strategy
        :return: List of ExcelTestDataRow
        """
        validate_result = {'is_pass': True, 'error_message': ''}
        # Parsing test data
        ws_test_data_rows = OrderedDict()
        for ws in self.parser_strategy.get_all_worksheet(wb):
            ws_column_indexs = self.parser_strategy.parsing_column_indexs(ws)

            validate_result = self.parser_strategy.is_ws_column_valid(ws, ws_column_indexs, validate_result)
            if validate_result['is_pass'] is not True:
                raise ValueError(validate_result['error_message'])

            test_data_rows = self.parser_strategy.parse_test_data_properties(ws, ws_column_indexs)
            ws_test_data_rows[ws.title] = test_data_rows
        print(str(datetime.now())+': Done parse data...')
        return ws_test_data_rows

    def insert_extra_columns(self, wb, columns):
        for ws in self.parser_strategy.get_all_worksheet(wb):
            ws_column_indexs = self.parser_strategy.parsing_column_indexs(ws)
            for column in reversed(columns):
                if column in ws_column_indexs:
                    continue
                ws.insert_cols(1)
                ws['A' + str(self.parser_strategy.start_row - 2)] = column
