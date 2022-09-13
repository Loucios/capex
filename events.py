import re
from openpyxl.utils import get_column_letter
from collections import Counter, defaultdict


class Events:
    def __init__(self,
                 energy_source_events: dict,
                 heating_network_events: dict,
                 tso_list: dict = None) -> None:
        self.energy_source_events = energy_source_events
        self.heating_network_events = heating_network_events
        self.tso_list = tso_list

    def create_base_source_events_table(self,
                                        calculations,
                                        workbook,
                                        table_names,
                                        header_table_names):
        # How much events
        events_number = (
            len(self.energy_source_events[header_table_names.number]) - 1
        )
        # How much events by tso
        tso_frequency = Counter(
            self.energy_source_events[header_table_names.tso_name]
        )
        # Dict for event number by tso
        num_tso_events = defaultdict(int)
        sum_tso_events = defaultdict(int)
        powers_length = len(header_table_names.power_index)
        power_tso_events = {
            tso: [0]*powers_length for tso in tso_frequency
            if tso is not None
        }
        year_sum_tso_events = {}
        # Run through every event
        for number in range(events_number):
            tso_name = (
                self.energy_source_events[header_table_names.tso_name][number]
            )
            sheet_name = self.get_short_name(tso_name)
            num_tso_events[tso_name] += 1
            # If tso first time
            if sheet_name not in workbook.sheetnames:
                # Create new sheet
                ws = workbook.create_sheet(title=sheet_name)
                # Insert main header
                ws = self.create_main_header(ws, tso_name, header_table_names)
                # Insert table header
                ws = self.create_base_source_table_header(
                    ws, calculations, table_names, header_table_names
                )
            # Insert events
            ws = workbook[sheet_name]
            ws, sum_tso_events = self.create_base_source_table_content(
                ws, calculations, header_table_names,
                num_tso_events[tso_name], number,
                sum_tso_events, year_sum_tso_events, power_tso_events, tso_name
            )
            # Insert footer
            if num_tso_events[tso_name] == tso_frequency[tso_name]:
                ws = self.create_base_source_table_footer(
                    ws, header_table_names, num_tso_events[tso_name] + 1,
                    tso_name, sum_tso_events, year_sum_tso_events,
                    power_tso_events, calculations, table_names
                )
        return workbook

    def create_main_header(self, ws, tso_name, header_table_names):
        row = header_table_names.main_header_row
        column = header_table_names.main_header_column
        ws.cell(row=row, column=column).value = tso_name
        style = header_table_names.base_source_table_head_style
        ws.cell(row=row, column=column).style = style
        return ws

    def create_base_source_table_header(self,
                                        ws,
                                        calculations,
                                        table_names,
                                        header_table_names):
        # Set column number
        column = header_table_names.table_header_column
        # Set row number
        row = header_table_names.table_header_row
        # Set header style
        style = header_table_names.base_source_table_head_table_style
        # Set height row
        height = header_table_names.base_source_table_head_height
        # Create header
        for structure in header_table_names.base_source_table_structure:
            if structure == header_table_names.head_capex_flow:
                # We must merge first row
                first_year = calculations.terms[table_names.first_year][0]
                last_year = calculations.terms[table_names.last_year][0]
                period = last_year - first_year + 1
                ws.merge_cells(
                    start_row=row, start_column=column,
                    end_row=row, end_column=column + period - 1
                )
                ws.cell(row=row, column=column).value = structure

                # And draw head_capex_flow cells with dates in second row
                for i in range(period):
                    ws.cell(row=row, column=column + i).style = style
                for year in range(first_year, last_year + 1):
                    ws.cell(row=row + 1, column=column).value = year
                    ws.cell(row=row + 1, column=column).style = style
                    column += 1
            else:
                # Base header cell
                ws.merge_cells(
                    start_row=row, start_column=column,
                    end_row=row + 1, end_column=column
                )
                ws.cell(row=row, column=column).value = structure
                ws.cell(row=row, column=column).style = style
                ws.cell(row=row + 1, column=column).style = style
                width = (
                    header_table_names.base_source_table_structure[structure]
                )
                ws.column_dimensions[get_column_letter(column)].width = width
                column += 1

        ws.row_dimensions[row + 1].height = height
        return ws

    def create_base_source_table_content(self,
                                         ws,
                                         calculations,
                                         header_table_names,
                                         serial_number,
                                         number,
                                         sum_tso_events,
                                         year_sum_tso_events,
                                         power_tso_events,
                                         tso_name):
        style = header_table_names.base_source_table_base_style
        column = 1
        row = serial_number + header_table_names.table_header_row + 1
        for structure in header_table_names.base_source_table_structure:

            if structure == header_table_names.number:
                # Number of order of this event in tso table
                ws.cell(row=row, column=column).value = serial_number

            elif structure == header_table_names.total_cost:
                # For this structure we need obtain capex
                unit_type = self.get_unit_type(header_table_names, number)
                power, index = self.get_power(header_table_names, number)
                # Save sum powers for footer
                if serial_number == 1:
                    powers_length = len(header_table_names.power_index)
                    power_tso_events[tso_name] = [0] * powers_length
                else:
                    power_tso_events[tso_name][index]
                power_tso_events[tso_name][index] += power

                capex = calculations.get_energy_source_capex(
                    power, unit_type, index
                )
                ws.cell(row=row, column=column).value = capex
                sum_tso_events[tso_name] += capex

            elif structure == header_table_names.head_capex_flow:
                # For this structure we need obtain capex flow
                terms = self.energy_source_events[
                                header_table_names.terms][number]
                capex_flow = calculations.get_capex_flow(
                    capex, terms, header_table_names.energy_sources
                )
                if serial_number == 1:
                    year_sum_tso_events[tso_name] = capex_flow
                for index, year in enumerate(capex_flow):
                    ws.cell(row=row, column=column).value = year
                    ws.cell(row=row, column=column).style = style
                    ws = self.set_number_format(
                        structure, ws, header_table_names, row, column
                    )
                    if serial_number > 1:
                        year_sum_tso_events[tso_name][index] += (
                            capex_flow[index]
                        )
                    column += 1
                column -= 1

            elif structure == header_table_names.total:
                # For this structure we need obtain sum of capex flow
                ws.cell(row=row, column=column).value = sum(capex_flow)

            else:
                # This structures goes unchanged
                value = self.energy_source_events[structure][number]
                ws.cell(row=row, column=column).value = value

            ws.cell(row=row, column=column).style = style
            ws = self.set_number_format(
                structure, ws, header_table_names, row, column
            )
            # Next column
            column += 1
        return ws, sum_tso_events

    def create_base_source_table_footer(self,
                                        ws,
                                        header_table_names,
                                        serial_number,
                                        tso_name,
                                        sum_tso_events,
                                        year_sum_tso_events,
                                        power_tso_events,
                                        calculations,
                                        table_names):
        style = header_table_names.base_source_table_base_style
        column = 1
        row = serial_number + header_table_names.table_header_row + 1
        for structure in header_table_names.base_source_table_structure:

            if structure == header_table_names.number:
                # Number of order of this event in tso table
                ws.cell(row=row, column=column).value = serial_number

            elif column == 2:
                value = header_table_names.total
                ws.cell(row=row, column=column).value = value

            elif structure in header_table_names.power_index:
                value = power_tso_events[tso_name][
                    header_table_names.power_index[structure]
                ]
                ws.cell(row=row, column=column).value = (
                    value if value else '-'
                )

            elif structure == header_table_names.total_cost:
                # For this structure we need obtain capex
                value = sum_tso_events.pop(tso_name, 0)
                ws.cell(row=row, column=column).value = value

            elif structure == header_table_names.head_capex_flow:
                # For this structure we need obtain capex flow
                for value in year_sum_tso_events[tso_name]:
                    ws.cell(row=row, column=column).value = value
                    ws.cell(row=row, column=column).style = style
                    ws = self.set_number_format(
                        structure, ws, header_table_names, row, column
                    )
                    column += 1
                column -= 1

            elif structure == header_table_names.terms:
                value = self.get_sum_terms(
                    year_sum_tso_events[tso_name], calculations, table_names)
                ws.cell(row=row, column=column).value = value

            elif structure == header_table_names.total:
                # For this structure we need obtain sum of capex flow
                value = sum(year_sum_tso_events[tso_name])
                ws.cell(row=row, column=column).value = value
                del year_sum_tso_events[tso_name]

            ws.cell(row=row, column=column).style = style
            ws = self.set_number_format(
                structure, ws, header_table_names, row, column
            )
            # Next column
            column += 1
        return ws

    def get_power(self,
                  header_table_names,
                  number):
        mw = self.energy_source_events[header_table_names.mw][number]
        th = self.energy_source_events[header_table_names.t_in_h][number]
        gh = self.energy_source_events[header_table_names.gcal_in_h][number]

        if mw != '-':
            power = mw
            index = header_table_names.power_index[header_table_names.mw]
        elif th != '-':
            power = 0.6 * th
            index = header_table_names.power_index[header_table_names.t_in_h]
        else:
            power = gh
            index = (
                header_table_names.power_index[header_table_names.gcal_in_h]
            )
        return power, index

    def get_unit_type(self, header_table_names, number):
        unit_type = header_table_names.enery_source_unit_type
        return self.energy_source_events[unit_type][number]

    def set_number_format(self,
                          structure,
                          ws,
                          header_table_names,
                          row,
                          column):
        mw = header_table_names.mw
        th = header_table_names.t_in_h
        gh = header_table_names.gcal_in_h
        total_cost = header_table_names.total_cost
        total = header_table_names.total
        capex_flow = header_table_names.head_capex_flow

        if structure == mw or structure == th:
            ws.cell(row=row, column=column).number_format = '# ##0'
        elif structure == gh:
            ws.cell(row=row, column=column).number_format = '# ##0.00'
        elif (structure == total_cost
              or structure == total
              or structure == capex_flow):
            ws.cell(row=row, column=column).number_format = '# ##0.0'

        return ws

    def get_short_name(self, tso_name):
        if self.tso_list is None:
            pattern = re.compile(r'[^\w ]+')
            return re.sub(pattern, '', tso_name)[:10]
        return self.tso_list[tso_name]

    def get_sum_terms(self, capex_flow, calculations, table_names):
        first_year = calculations.terms[table_names.first_year][0]
        last_year = calculations.terms[table_names.last_year][0]
        start_year = 0
        end_year = first_year
        for year in range(last_year - first_year + 1):
            if capex_flow[year]:
                if not start_year:
                    start_year = first_year + year
                end_year = first_year + year
        if not start_year:
            return ''
        if start_year == end_year:
            return start_year
        return f'{start_year} - {end_year}'
