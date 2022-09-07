class Calculations:
    def __init__(self, dict1: dict, dict2: dict) -> None:
        """
        energy_source_unit_costs = {
            'powers': [list of powers],
            'type 1': [list of specific type 1 unit costs],
            'type 2': [list of specific type 2 unit costs],
            ...
        }
        heating_network_unit_costs = {
            'diameters': [list of diameters],
            'layer_type_name & constructions_type_1':
                [list of specific layer type and construction type unit costs],
            'layer_type_name & constructions_type_2':
                [list of specific layer type and construction type unit costs],
            ...
        }
        """
        self.energy_source_unit_costs = dict1
        self.heating_network_unit_costs = dict2

    def get_energy_source_capex(self, power: float, unit_type: str) -> float:
        powers = self.energy_source_unit_costs.get('Диапазон мощности')
        # we need to find the nearest power for our value
        # then define the unit cost corresponds this power
        unit_costs = self.energy_source_unit_costs.get(unit_type)
        unit_cost = unit_costs[self.binary_search(power, powers)]
        # then define the construction work cost for the object with this power
        return power * unit_cost

    def get_heating_network_cost(self,
                                 diameter: float,
                                 length: float,
                                 laying_type: str,
                                 unit_type: str) -> float:
        diameters = self.heating_network_unit_costs.get('diameter')
        unit_costs = self.heating_network_unit_costs.get(laying_type
                                                         + unit_type)
        unit_cost = unit_costs[self.binary_search(diameter, diameters)]
        return length * unit_cost

    def binary_search(self, value: float, array: list) -> int:
        right = len(array) - 1
        left = 0
        while right - left != 1:
            center = (left + right) // 2
            if array[center] == value:
                return center
            elif value > array[center]:
                left = center
            else:
                right = center
        return left if array[right] - value > value - array[left] else right

    def get_capex_flow(self,
                       capex: float,
                       terms: dict,
                       deadlines: dict,
                       deflators: dict,
                       design_rate: float) -> dict:
        # create capex flow
        capex_flow = {}
        for year in range(terms['end'] - terms['start'] + 1):
            capex_flow[terms['start'] + year] = 0
        # define deflator for first year - 1 of construction work
        deflator = 1
        for year in range(deadlines['start'] - terms['prices_year']):
            deflator *= deflators.get(str(terms['prices_year'] + year))
        # fill the capex flow
        # in the first year we carry out design and survey work
        time = deadlines['end'] - deadlines['start'] + 1
        for year in range(time):
            deflator *= deflators.get(str(deadlines['start'] + year))
            if time == 1:
                capex_flow[deadlines['start']] = capex * deflator
            elif year:
                capex_flow[
                    deadlines['start'] + year
                ] = capex * (1 - design_rate) * deflator / (time - 1)
            else:
                capex_flow[deadlines['start']] = capex * design_rate * deflator
        return capex_flow
