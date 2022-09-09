from openpyxl import load_workbook
from calculations import Calculations
from events import Events
from names import Names
from styles import head_style, base_style, head_table_style
from tables import Tables


def main():
    names = Names()
    wb_for_data = Tables(
        load_workbook(filename=names.filename, data_only=True)
    )
    calculations = Calculations(
        wb_for_data.get_table_data(names.energy_source_unit_costs),
        wb_for_data.get_table_data(names.heating_network_unit_costs),
        wb_for_data.get_table_data(names.deflators),
        wb_for_data.get_table_data(names.terms),
        wb_for_data.get_small_table_data(names.stages),
        wb_for_data.get_table_data(names.nds),
    )
    events = Events(
        wb_for_data.get_table_data(names.energy_source_events),
        wb_for_data.get_table_data(names.heating_network_events),
    )
    wb_for_write = load_workbook(filename=names.filename, keep_vba=True)
    wb_for_write.add_named_style(base_style)
    wb_for_write.add_named_style(head_style)
    wb_for_write.add_named_style(head_table_style)
    wb_for_write = events.create_base_source_events_table(
        calculations, wb_for_write
    )
    wb_for_write.save('new_' + names.filename)


if __name__ == '__main__':
    main()
