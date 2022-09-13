from dataclasses import dataclass

'''
@dataclass(frozen=True)
class Titles:
    number = '№ п/п'
    tso_name = 'Наименование EТО'
    eto_name = 'Наименование ТСО'
    district = 'Aдминистративный район'
    source_type = 'ТЭЦ/ котельная'
    id_source = 'id источника'
    source = 'Наименование источника'
    event_title = 'Наименование мероприятия'
    terms = 'Cроки'
    base_direction = 'Основное направление'
    unit_type = 'Тип мероприятия'
    laying_type = 'Типа прокладки'
    length = 'Протяженность, м'
    diameter = 'Диаметр, мм'
    power = 'Гкал/ч'
    desing = 'ПИР'
    inv_pro = 'ИП'
    chapter_8_directions = 'Направления мероприятий Глава 8'
    chapter_12_directions = 'Направления мероприятий из Главы 12'
    total_cost = 'Общая стоимость мероприятий, млн руб. без НДС'
    total = 'Итого'
    mw = 'МВт'
    gh = 'Гкал/ч'
    th = 'т/ч'
    base_style = 'base_style'


class StyleMixin:
    @property
    def base_style(cls):
        return 'base_style'

    @property
    def title_style(cls):
        return 'base_style'

    @property
    def header_style(cls):
        return 'header_style'

    @property
    def header_height(cls):
        return 70

    @property
    def header_position(cls):
        return {
            'title_row': 1,
            'title_column': 1,
            'header_row': 3,
            'header_column': 1,
        }



        


@dataclass
class NetworkEventsBaseTableTitles:
    number: str = Titles.number
    district: str = Titles.district
    source: str = Titles.source
    event_title: str = Titles.event_title
    terms: str = Titles.terms
    laying_type: str = Titles.laying_type
    length: str = Titles.length
    diametr: str = Titles.diameter
    power: str = Titles.power
    total_cost: str = total_cost
    total: str = Titles.total


@dataclass
class SourceEventsBaseTableTitles:
    number: str = Titles.number
    source: str = Titles.source
    event_title: str = Titles.event_title
    terms: str = Titles.terms
    mw: str = Titles.mw
    gh: str = Titles.gh
    th: str = Titles.th
    total_cost: str = Titles.total_cost
    total: str = Titles.total


@dataclass(frozen=True)
class AuxiliaryTable:
    title = 'Вспомогательные'
    unexpected_expenses = 'Непредвидимые расход'
    tfu_unit_cost = ''


@dataclass(frozen=True)
class NetworkEventsTitles:
    title = 'МероприятияСети'
    base_table_structure = []


class WidthsMixin:
    widths = {
        Titles.number: 
        Titles.district
        Titles.source
        Titles.event_title
        Titles.terms
        Titles.laying_type
        Titles.length
        Titles.diameter
        Titles.power
        Titles.total_cost
        Titles.total
    }
    
    @property
    def widths(cls):

'''


@dataclass(frozen=True)
class TableNames:
    filename = 'capex.xlsm'
    energy_source_unit_costs = 'УдельникиИсточники'
    heating_network_unit_costs = 'УдельникиСети'
    tfu_unit_cost = 'ТФУ'
    deflators = 'Индексы'
    terms = 'Сроки'
    stages = 'Этапы'
    nds = 'НДС'
    energy_source_events = 'МероприятияИсточники'
    heating_network_events = 'МероприятияСети'
    first_year = 'Год начала'
    last_year = 'Год окончания'
    tso_list = 'СписокТСО'


@dataclass(frozen=True)
class HeaderTableNames:
    number = '№ п/п'
    tso_name = 'Наименование ТСО'
    eto_name = 'Наименование ЕТО'
    event_name = 'Наименование мероприятия'
    mw = 'МВт'
    gcal_in_h = 'Гкал/ч'
    t_in_h = 'т/ч'
    terms = 'Сроки'
    total_cost = 'Общая стоимость мероприятий, млн руб. без НДС'
    total = 'Итого'
    head_capex_flow = (
        'Капитальные затраты в прогнозных ценах с учетом НДС, млн руб.'
    )
    power_index = {
        mw: 0,
        gcal_in_h: 1,
        t_in_h: 2
    }
    energy_sources = 'Источники тепловой энергии'
    # The heat source event table we create
    base_source_table_structure = {
        number: 6,
        # tso_name: 16,
        # 'Административный район': 21,
        'Наименование источника': 16,
        event_name: 16,
        terms: 10,
        mw: 7,
        gcal_in_h: 9,
        t_in_h: 7,
        total_cost: 14,
        head_capex_flow: 0,
        total: 10,
    }
    # Styles
    base_source_table_base_style = 'base_style'
    base_source_table_head_style = 'head_style'
    base_source_table_head_table_style = 'head_table_style'
    base_source_table_head_height = 70
    # Headers
    main_header_row = 1
    main_header_column = 1
    table_header_row = 3
    table_header_column = 1
    # The source table
    enery_source_unit_type = 'Тип мероприятия'
