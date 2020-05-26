from openpyxl.utils import column_index_from_string
from openpyxl.utils.cell import coordinate_from_string
from ExcelDataDriver.ExcelParser.ABCParserStrategy import ABCParserStrategy
from ExcelDataDriver.ExcelTestDataRow import ExcelTestDataRow


class CustomExcelParser(ABCParserStrategy):
    
    def __init__(self, main_column_key=None):
        ABCParserStrategy.__init__(self, main_column_key)
        print('Using CustomExcelParser')
        self.maximum_column_index_row = 2
        self.start_row = 3

    def parsing_column_indexs(self, ws):
        '''
        @param sku: Product SKU
        @param normal price: Product normal price
        @param normal cost in vat: Product normal cost include vat
        @param normal cost ex vat: Product normal cost exclude vat
        @param normal cost gp: Product normal GP %
        @param promotion price: Product promotion price
        @param promotion cost in vat: Product promotion cost include vat
        @param promotion cost ex vat: Product promotion cost exclude vat
        @param promotion cost gp: Product promotion GP %
        '''
        ws_column_indexs = {}
        
        ##########################
        # Parse mandatory property
        ##########################
        for index, row in enumerate(ws.rows):
            if index > self.maximum_column_index_row:
                break
            for cell in row:
                if (cell.value is not None) and (cell.value in self.DEFAULT_COLUMN_INDEXS):
                    ws_column_indexs[cell.value] = column_index_from_string(coordinate_from_string(cell.coordinate)[0])
                    print('Mandatory : '+str(cell.value) + ' : ' + str(cell.coordinate) + ' : ' + str(column_index_from_string(coordinate_from_string(cell.coordinate)[0])))
            if len(ws_column_indexs) > 0:
                break
            
        ##########################
        # Parse optional property
        ##########################
        #   Normal Price
        #     Normal Price
        #     (In Vat)
        #     GP (%)
        #   Promotion Price
        #     Promotion Price
        #     (In Vat)
        #     GP (%)
        for index, row in enumerate(ws.rows):
            if index > self.maximum_column_index_row:
                break
            for cell in row:
                if (cell.value is not None) and (cell.value not in self.DEFAULT_COLUMN_INDEXS):
                    column_name = cell.value.lower().strip()
                    
                    # Normal & Promotion
                    if column_name == '(in vat)':
                        if 'normal cost in vat' not in ws_column_indexs:
                            ws_column_indexs['normal cost in vat'] = column_index_from_string(coordinate_from_string(cell.coordinate)[0])
                        else:
                            ws_column_indexs['promotion cost in vat'] = column_index_from_string(coordinate_from_string(cell.coordinate)[0])
                    
                    elif column_name == '(ex vat)':
                        if 'normal cost ex vat' not in ws_column_indexs:
                            ws_column_indexs['normal cost ex vat'] = column_index_from_string(coordinate_from_string(cell.coordinate)[0])
                        else:
                            ws_column_indexs['promotion cost ex vat'] = column_index_from_string(coordinate_from_string(cell.coordinate)[0])
                    
                    elif column_name == 'gp (%)':
                        if 'normal cost gp' not in ws_column_indexs:
                            ws_column_indexs['normal cost gp'] = column_index_from_string(coordinate_from_string(cell.coordinate)[0])
                        else:
                            ws_column_indexs['promotion cost gp'] = column_index_from_string(coordinate_from_string(cell.coordinate)[0])

                    ws_column_indexs[column_name] = column_index_from_string(coordinate_from_string(cell.coordinate)[0])
                    print('Optional : '+str(column_name) + ' : ' + str(cell.coordinate) + ' : ' + str(column_index_from_string(coordinate_from_string(cell.coordinate)[0])))
        
        print('Done parsing column indexes')
        return ws_column_indexs
    
    def parse_test_data_properties(self, ws, ws_column_indexs):
        test_datas = []
        for index, row in enumerate(ws.rows):
            if index < self.start_row:
                continue
            self.is_test_data_valid(ws_column_indexs, ws.title, index, row)
            test_data = self.map_data_row_into_test_data_obj(ws_column_indexs, ws.title, index, row)
            if test_data is not None:
                test_datas.append(test_data)
            else:
                break
        print('Total test datas: ' + str(len(test_datas)))
        return test_datas

    def is_test_data_valid(self, ws_column_indexes, ws_title, row_index, row):
        return True

    def map_data_row_into_test_data_obj(self, ws_column_indexes, ws_title, row_index, row):
        status = row[ws_column_indexes[self.MANDATORY_TEST_DATA_COLUMN['status']] - 1]
        log_message = row[ws_column_indexes[self.MANDATORY_TEST_DATA_COLUMN['log_message']] - 1]
        screenshot = row[ws_column_indexes[self.MANDATORY_TEST_DATA_COLUMN['screenshot']] - 1]
        tags = row[ws_column_indexes[self.MANDATORY_TEST_DATA_COLUMN['tags']] - 1]
        koriico_sku = row[ws_column_indexes['sku'] - 1]
        
        # Excel library send the last row with None data.
        if koriico_sku.value is None:
            return None
        test_data_row = ExcelTestDataRow(ws_title, row_index, row, ws_column_indexes, status, log_message, screenshot, tags)
        return test_data_row
