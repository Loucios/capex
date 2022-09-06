# from openpyxl import load_workbook
from calculations import Calculations


def main():
    table1 = {}
    table1['power'] = [2, 5, 10, 13, 20, 25, 50]
    table1['new'] = [30, 25, 20, 15, 10, 5, 2]

    table2 = {}
    table2['diameter'] = [80, 100, 125, 150]
    table2['канальная2'] = [1, 2, 3, 4]
    table2['канальная1'] = [5, 6, 7, 8]
    calculations = Calculations(table1, table2)
    energy_source_capex = calculations.get_energy_source_capex(50, 'new')
    heating_network_capex = calculations.get_heating_network_cost(
        100, 10, 'канальная', '1'
    )

    terms = {
        'prices_year': 2021,
        'start': 2022,
        'end': 2025
    }
    deadlines = {
        'start': 2022,
        'end': 2023
    }
    deflators = {
        '2020': 1,
        '2021': 1.1,
        '2022': 1.1,
        '2023': 1.1,
        '2024': 1.1,
        '2025': 1.1
    }
    design = 0.1
    energy_source_capex_flow = calculations.get_capex_flow(
        energy_source_capex, terms, deadlines, deflators, design
    )
    heating_network_capex_flow = calculations.get_capex_flow(
        heating_network_capex, terms, deadlines, deflators, design
    )
    print(energy_source_capex_flow)
    print(heating_network_capex_flow)


if __name__ == '__main__':
    main()
