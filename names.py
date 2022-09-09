from dataclasses import dataclass


@dataclass
class Names:
    filename: str = 'capex.xlsm'
    energy_source_unit_costs: str = 'УдельникиИсточники'
    heating_network_unit_costs: str = 'УдельникиСети'
    deflators: str = 'Индексы'
    terms: str = 'Сроки'
    stages: str = 'Этапы'
    nds: str = 'НДС'
    energy_source_events: str = 'МероприятияИсточники'
    heating_network_events: str = 'МероприятияСети'
