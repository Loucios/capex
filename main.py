from openpyxl import load_workbook
from calculations import Calculations
from events import Events
from names import HeaderTableNames, TableNames
from styles import head_style, base_style, head_table_style
from tables import Tables


def main():
    table_names = TableNames()
    filename = table_names.filename
    header_table_names = HeaderTableNames()
    wb_for_data = Tables(
        load_workbook(filename=filename, data_only=True)
    )
    calculations = Calculations(
        wb_for_data.get_table_data(table_names.energy_source_unit_costs),
        wb_for_data.get_table_data(table_names.heating_network_unit_costs),
        wb_for_data.get_table_data(table_names.tfu_unit_cost),
        wb_for_data.get_table_data(table_names.deflators),
        wb_for_data.get_table_data(table_names.terms),
        wb_for_data.get_small_table_data(table_names.stages),
        wb_for_data.get_table_data(table_names.nds),
    )
    events = Events(
        wb_for_data.get_table_data(table_names.energy_source_events),
        wb_for_data.get_table_data(table_names.heating_network_events),
        wb_for_data.get_little_table_data(table_names.tso_list),
    )
    wb_for_write = load_workbook(filename=filename, keep_vba=True)
    wb_for_write.add_named_style(base_style)
    wb_for_write.add_named_style(head_style)
    wb_for_write.add_named_style(head_table_style)
    wb_for_write = events.create_base_source_events_table(
        calculations, wb_for_write, table_names, header_table_names
    )
    wb_for_write.save('new_' + filename)


if __name__ == '__main__':
    main()
