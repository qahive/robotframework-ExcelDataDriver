from ExcelDataDriver.Config.DataTypes import DataTypes


class ExcelReferenceDataRow(object):

    def __init__(self, excel_title, excel_row_index, excel_row, column_indexes):
        self.excel_title = excel_title
        self.excel_row_index = excel_row_index
        self.row_no = None
        self.sheet_name = None
        self.properties_list = dict()

        for column_index_key in column_indexes:
            if self.row_no is None:
                self.row_no = excel_row[column_indexes[column_index_key] - 1].row
                self.sheet_name = excel_row[column_indexes[column_index_key] - 1].parent.title
            self.properties_list[column_index_key.lower().strip().replace(' ', '_')] = excel_row[column_indexes[column_index_key] - 1].value

    def get_data_type(self):
        return DataTypes.REFERENCE_DATA

    def get_row_no(self):
        return self.row_no

    def get_sheet_name(self):
        return self.sheet_name

    def get_test_data_property(self, property_name):
        try:
            return self.properties_list[property_name.lower().strip()]
        except:
            raise Exception('Can\'t find property name '+property_name+' under test data row index '+str(self.get_row_no()))
