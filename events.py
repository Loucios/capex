import re
from openpyxl.utils import get_column_letter


class Events:
    def __init__(self,
                 energy_source_events,
                 heating_network_events) -> None:
        self.energy_source_events = energy_source_events
        self.heating_network_events = heating_network_events
        self.energy_source_tsos = set(
            self.energy_source_events['Наименование ТСО']
        )
        self.heating_network_tsos = set(
            self.heating_network_events['Наименование ТСО']
        )
        self.table_structure = [
            '№ п/п',
            'Наименование ТСО',
            # 'Наименование ЕТО',
            'Административный район',
            'Наименование источника',
            'Наименование мероприятия',
            'Сроки',
            'МВт',
            'Гкал/ч',
            'т/ч',
            'Общая стоимость мероприятий, млн руб. без НДС',
        ]
        self.widths = [6, 16, 21, 16, 16, 10, 7, 7, 7, 14]
        self.base_style = 'base_style'
        self.head_style = 'head_style'
        self.head_table_style = 'head_table_style'

    def create_base_source_events_table(self, calculations, workbook):
        pattern = re.compile(r'[^\w ]+')
        events_number = len(self.energy_source_events['№ п/п']) - 1
        temp = {
            tso: [1, re.sub(pattern, '', tso)[:15]]
            for tso in self.energy_source_tsos
            if tso is not None
        }
        for number in range(events_number):
            tso_name = self.energy_source_events['Наименование ТСО'][number]
            serial_number, sheet_name = temp[tso_name]
            # Create table header
            if sheet_name not in workbook.sheetnames:
                ws = workbook.create_sheet(title=sheet_name)
                index = 1
                ws.cell(row=1, column=index).value = tso_name
                ws.cell(row=1, column=index).style = self.head_style
                for structure in self.table_structure:
                    ws.merge_cells(
                        start_row=3, start_column=index,
                        end_row=4, end_column=index
                    )
                    ws.cell(row=3, column=index).value = structure
                    ws.cell(row=3, column=index).style = self.head_table_style
                    ws.cell(row=4, column=index).style = self.head_table_style
                    ws.column_dimensions[get_column_letter(index)].width = (
                        self.widths[index - 1]
                    )
                    index += 1
                period_start = calculations.terms['Год начала'][0]
                period_end = calculations.terms['Год окончания'][0]
                period = period_end - period_start + 1
                ws.merge_cells(
                    start_row=3, start_column=index,
                    end_row=3, end_column=index + period - 1
                )
                ws.cell(row=3, column=index).value = (
                    'Капитальные затраты в прогнозных ценах с учетом НДС, ' +
                    'млн руб.'
                )
                for i in range(period):
                    ws.cell(row=3, column=index + i).style = (
                        self.head_table_style
                    )
                for year in range(period_start, period_end + 1):
                    ws.cell(row=4, column=index).value = year
                    ws.cell(row=4, column=index).style = self.head_table_style
                    index += 1
                ws.merge_cells(
                        start_row=3, start_column=index,
                        end_row=4, end_column=index
                    )
                ws.cell(row=3, column=index).value = 'Итого'
                ws.cell(row=3, column=index).style = self.head_table_style
                ws.cell(row=4, column=index).style = self.head_table_style
                ws.row_dimensions[4].height = 70

            # Insert events
            ws = workbook[sheet_name]
            index = 1
            for structure in self.table_structure:

                if structure == '№ п/п':
                    ws.cell(row=serial_number + 4, column=index).value = (
                        serial_number
                    )

                elif index == len(self.table_structure):
                    # Obtain power
                    unit_type = (
                        self.energy_source_events['Тип мероприятия'][number]
                    )
                    if (
                        isinstance(
                            self.energy_source_events['МВт'][number], float
                        ) or
                        isinstance(
                            self.energy_source_events['МВт'][number], int
                        )
                    ):
                        power = self.energy_source_events['МВт'][number] * 10
                    elif (
                        isinstance(
                            self.energy_source_events['т/ч'][number], float
                        ) or
                        isinstance(
                            self.energy_source_events['т/ч'][number], int
                        )
                    ):
                        power = (
                                0.6 * self.energy_source_events['т/ч'][number]
                            )
                    else:
                        power = self.energy_source_events['Гкал/ч'][number]

                    # Obtain capex
                    capex = calculations.get_energy_source_capex(
                        power, unit_type
                    )
                    ws.cell(row=serial_number + 4, column=index).value = (
                        capex
                    )

                else:
                    ws.cell(row=serial_number + 4, column=index).value = (
                        self.energy_source_events[structure][number]
                    )

                ws.cell(row=serial_number + 4, column=index).style = (
                    self.base_style
                )

                if structure == 'МВт' or structure == 'т/ч':
                    ws.cell(
                        row=serial_number + 4, column=index
                    ).number_format = '0'
                elif structure == 'Гкал/ч':
                    # print(1)
                    ws.cell(
                        row=serial_number + 4, column=index
                    ).number_format = '0.00'
                elif structure == 'Общая стоимость мероприятий, млн руб. без НДС':
                    ws.cell(
                        row=serial_number + 4, column=index
                    ).number_format = '0.0'

                index += 1

            time = str(self.energy_source_events['Сроки'][number])
            if time[:4] == time[-4:]:
                start = end = int(time)
            else:
                start = int(time[:4])
                end = int(time[-4:])
            capex_flow = calculations.get_capex_flow(
                capex, start, end, 'Источники тепловой энергии'
            )

            for year in capex_flow:
                ws.cell(row=serial_number + 4, column=index).value = year
                ws.cell(row=serial_number + 4, column=index).style = (
                    self.base_style
                )
                ws.cell(
                    row=serial_number + 4, column=index
                ).number_format = '# ##0.0'
                index += 1

            ws.cell(row=serial_number + 4, column=index).value = (
                sum(capex_flow)
            )
            ws.cell(row=serial_number + 4, column=index).style = (
                    self.base_style
                )
            ws.cell(
                row=serial_number + 4, column=index
            ).number_format = '# ##0.0'

            temp[tso_name][0] += 1

        return workbook
