class Tables:
    def __init__(self, workbook) -> None:
        self.workbook = workbook
        self.named_tables = {}

        for worksheet in self.workbook.sheetnames:
            for table in self.workbook[worksheet].tables:
                self.named_tables[table] = worksheet

    def get_table_data(self, table_name: str) -> dict:
        worksheet = self.workbook[self.named_tables[table_name]]
        table_range = worksheet.tables[table_name].ref

        table_head = worksheet[table_range][0]
        table_data = worksheet[table_range][1:]

        columns = [column.value for column in table_head]
        data = {column: [] for column in columns}

        for row in table_data:
            row_value = [cell.value for cell in row]
            for key, value in zip(columns, row_value):
                data[key].append(value)

        return data

    def get_small_table_data(self, table_name: str) -> dict:
        worksheet = self.workbook[self.named_tables[table_name]]
        table_range = worksheet.tables[table_name].ref

        table_head = worksheet[table_range][0]
        table_data = worksheet[table_range][1:]

        data = {}
        for row in table_data:
            for index, cell in enumerate(row):
                if index == 1:
                    data[cell.value] = {}
                elif index > 1:
                    data[row[1].value][table_head[index].value] = cell.value
        return data
