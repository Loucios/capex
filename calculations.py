class Calculations:
    def __init__(self, unit_costs: dict) -> None:
        """
        unit_costs = {
            'powers': [list of powers],
            'type 1': [list of specific type 1 unit costs],
            'type 2': [list of specific type 2 unit costs]
        }
            type 1 construction works unit cost (unit_costs[type_1][i])
            corresponds specific construction object power (powers[i])
        """
        self.unit_costs = unit_costs

    def get_energy_source_capex(self, power: float, unit_type: str) -> float:
        powers = self.unit_costs.get('power')
        # we need to find the nearest power for our value
        # then define the unit cost corresponds this power
        unit_costs = self.unit_costs.get(unit_type)
        unit_cost = unit_costs[self.binary_search(power, powers)]
        # then define the construction work cost for the object with this power
        return power * unit_cost

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
                       design: float) -> dict:
        # create capex flow
        capex_flow = {}
        for year in range(terms['end'] - terms['start'] + 1):
            capex_flow[terms['start'] + year] = 0
            # print(capex_flow)
        # define deflator for first year - 1 of construction work
        deflator = 1
        for year in range(deadlines['start'] - terms['prices_year']):
            '''print(deflators)
            print(terms['prices_year'] + year)
            print(deflators.get(str(terms['prices_year'] + year)))'''
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
                ] = capex * (1 - design) * deflator / (time - 1)
            else:
                capex_flow[deadlines['start']] = capex * design * deflator
        return capex_flow
