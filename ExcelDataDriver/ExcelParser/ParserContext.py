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

        # Validate column
        print('Validate reference column')
        for ws in self.parser_strategy.get_all_worksheet(wb):
            validate_result = self.parser_strategy.is_ws_column_valid(ws, validate_result)
        if validate_result['is_pass'] is not True:
            raise ValueError(validate_result['error_message'])
        print('Done validate reference column')

        # Parsing test data
        print('Parsing reference data')
        ws_test_data_rows = OrderedDict()
        for ws in self.parser_strategy.get_all_worksheet(wb):
            ws_column_indexs = self.parser_strategy.parsing_column_indexs(ws)
            test_data_rows = self.parser_strategy.parse_test_data_properties(ws, ws_column_indexs)
            ws_test_data_rows[ws.title] = test_data_rows
        print('Done validate reference data')

        return ws_test_data_rows
