class TablesByName:
    def __init__(self, workbook) -> None:
        self.workbook = workbook
        self.named_tables = {}

        for worksheet in self.workbook.sheetnames:
            for table in self.workbook[worksheet].tables:
                self.named_tables[table] = worksheet

    def get_named_tables(self):
        return self.named_tables

    def get_table_data(self, table_name):
        worksheet = self.workbook[self.named_tables[table_name]]
        table_range = worksheet.tables[table_name].ref

        table_head = worksheet[table_range][0]
        table_data = worksheet[table_range][1:]

        columns = [column.value for column in table_head]
        data = {column: [] for column in columns}

        for row in table_data:
            row_val = [cell.value for cell in row]
            for key, val in zip(columns, row_val):
                data[key].append(val)

        return data
