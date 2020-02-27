from ExcelDataDriver.ExcelReferenceDataRow.ExcelReferenceDataRow import ExcelReferenceDataRow
from ExcelDataDriver.ExcelParser.ABCParserStrategy import ABCParserStrategy


class DefaultReferenceParserStrategy(ABCParserStrategy):

    def __init__(self):
        ABCParserStrategy.__init__(self)

    def is_ws_column_valid(self, ws, validate_result):
        return validate_result

    def is_test_data_valid(self, ws_column_indexes, ws_title, row_index, row):
        return True

    def map_data_row_into_test_data_obj(self, ws_column_indexes, ws_title, row_index, row):
        # Excel library send the last row with None data.
        if row[-1].value is None:
            return None

        test_data_row = ExcelReferenceDataRow(ws_title, row_index, row, ws_column_indexes)
        return test_data_row
