class Calculations:
    def __init__(self,
                 energy_source_unit_costs: dict,
                 heating_network_unit_costs: dict,
                 tfu_unit_cost: dict,
                 deflators: dict,
                 terms: dict,
                 stages: dict,
                 nds: dict) -> None:
        self.energy_source_unit_costs = energy_source_unit_costs
        self.heating_network_unit_costs = heating_network_unit_costs
        self.tfu_unit_cost = tfu_unit_cost
        self.deflators = deflators
        self.terms = terms
        self.stages = stages
        self.nds = nds

    def get_energy_source_capex(self,
                                power: float,
                                unit_type: str,
                                is_tfu: bool,
                                titles) -> float:

        if is_tfu:
            unit_cost = self.tfu_unit_cost.get(unit_type)[0]
        else:
            powers = self.energy_source_unit_costs.get(titles.power_range)
            unit_costs = self.energy_source_unit_costs.get(unit_type)
            unit_cost = unit_costs[self.binary_search(power, powers)]

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
                       time: str,
                       otype: str,
                       titles) -> dict:
        time = str(time)
        if time[:4] == time[-4:]:
            start = end = int(time)
        else:
            start = int(time[:4])
            end = int(time[-4:])
        # create capex flow
        capex_flow = []
        time = end - start + 1
        deflator = 1
        design_rate = self.stages[titles.design][otype] / 100
        for index, year in enumerate(self.deflators[titles.year]):
            if year >= self.terms[titles.year_cost][0]:
                deflator *= self.deflators[titles.deflator][index]

            # fill the capex flow
            if (
                self.terms[titles.first_year][0] <= year
                <= self.terms[titles.last_year][0]
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
                    capex_flow[-1] *= (1 + self.nds[titles.nds][0])
                else:
                    capex_flow.append(0)

        return capex_flow
