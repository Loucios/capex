# from openpyxl import load_workbook
from calculations import Calculations


def main():
    table = {}
    table['power'] = [2, 5, 10, 13, 20, 25, 50]
    table['new'] = [30, 25, 20, 15, 10, 5, 2]

    calculations = Calculations(table)
    capex = calculations.get_energy_source_capex(50, 'new')

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
        '2023': 1,
        '2024': 1,
        '2025': 1
    }
    design = 0.1
    capex_flow = calculations.get_capex_flow(
        capex, terms, deadlines, deflators, design
    )
    print(capex_flow)


if __name__ == '__main__':
    main()
