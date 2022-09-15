from events import SourceEvents
from names import SourceEventsBaseTableTitles, Titles
from openpyxl import load_workbook
from styles import Style
from tables import OriginTables


def main():
    titles = Titles()

    wb = load_workbook(filename=titles.filename, keep_vba=True)
    style = Style()
    wb = style.add_styles(wb)

    tables = OriginTables(titles)
    table = SourceEventsBaseTableTitles()
    events = SourceEvents(tables.energy_source_events, tables.tso_list)
    wb = events.create_table(tables, wb, titles, table)
    wb.save('new_' + titles.filename)


if __name__ == '__main__':
    main()
