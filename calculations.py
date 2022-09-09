class Calculations:
    def __init__(self,
                 energy_source_unit_costs: dict,
                 heating_network_unit_costs: dict,
                 deflators: dict,
                 terms: dict,
                 stages: dict,
                 nds: dict) -> None:
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
        deflators = {
            '№ п/п': [],
            'Год': [list of years],
            'Индекс': [list of deflators],
        }
        """
        self.energy_source_unit_costs = energy_source_unit_costs
        self.heating_network_unit_costs = heating_network_unit_costs
        self.deflators = deflators
        self.terms = terms
        self.stages = stages
        self.nds = nds

    def get_energy_source_capex(self, power: float, unit_type: str) -> float:
        powers = self.energy_source_unit_costs.get('Диапазон мощности')
        # we need to find the nearest power for our value
        # then define the unit cost corresponds this power
        unit_costs = self.energy_source_unit_costs.get(unit_type)
        unit_cost = unit_costs[self.binary_search(power, powers)]
        # then define the construction work cost for the object with this power
        return power * unit_cost

    def get_heating_network_capex(self,
                                  diameter: float,
                                  length: float,
                                  laying_type: str,
                                  unit_type: str) -> float:
        diameters = self.heating_network_unit_costs.get('2Ду, мм')
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
                       start: int,
                       end: int,
                       object_type: str) -> dict:
        # create capex flow
        capex_flow = []
        time = end - start + 1
        deflator = 1
        design_rate = self.stages['ПИР'][object_type] / 100
        for index, year in enumerate(self.deflators['Год']):
            if year >= self.terms['Цены, год'][0]:
                deflator *= self.deflators['Индекс'][index]
            # fill the capex flow
            if (
                self.terms['Год начала'][0] <= year
                <= self.terms['Год окончания'][0]
            ):
                if start <= year <= end:
                    if time == 1:
                        capex_flow.append(capex * deflator)
                    elif year - start:
                        capex_flow.append(
                            capex * (1 - design_rate) * deflator / (time - 1)
                        )
                    else:
                        # in the first year we carry out design and survey work
                        capex_flow.append(capex * design_rate * deflator)
                    capex_flow[-1] *= (1 + self.nds['НДС'][0])
                else:
                    capex_flow.append(0)

        return capex_flow
