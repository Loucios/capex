from dataclasses import dataclass


@dataclass(frozen=True)
class Titles:
    base_direction = 'Основное направление'
    base_style = 'base_style'
    capex_flow = (
        'Капитальные затраты в прогнозных ценах с учетом НДС, млн руб.'
    )
    chapter_12_directions = 'Направления мероприятий из Главы 12'
    chapter_8_directions = 'Направления мероприятий Глава 8'
    chp_unit_costs = 'ЦТПУдельники'
    deflator = 'Индекс'
    deflators = 'Индексы'
    design = 'ПИР'
    diameter = 'Диаметр, мм'
    district = 'Административный район'
    energy_source_events = 'МероприятияИсточники'
    energy_source_unit_costs = 'УдельникиИсточники'
    energy_sources = 'Источники тепловой энергии'
    eto_name = 'Наименование EТО'
    event_title = 'Наименование мероприятия'
    filename = 'capex.xlsm'
    first_year = 'Год начала'
    footer_style = 'footer_style'
    gh = 'Гкал/ч'
    header_style = 'header_style'
    heating_network_events = 'МероприятияСети'
    heating_network_unit_costs = 'УдельникиСети'
    id_source = 'id источника'
    inv_pro = 'ИП'
    last_year = 'Год окончания'
    laying_type = 'Типа прокладки'
    length = 'Протяженность, м'
    mw = 'МВт'
    nds = 'НДС'
    number = '№ п/п'
    power = 'Гкал/ч'
    power_range = 'Диапазон мощности'
    source = 'Наименование источника'
    source_type = 'ТЭЦ/ котельная'
    stages = 'Этапы'
    terms = 'Сроки'
    tfu_unit_cost = 'ТФУ'
    th = 'т/ч'
    title_style = 'title_style'
    total = 'Итого'
    total_cost = 'Общая стоимость мероприятий, млн руб. без НДС'
    tso_list = 'СписокТСО'
    tso_name = 'Наименование ТСО'
    unit_type = 'Тип мероприятия'
    year = 'Год'
    year_cost = 'Цены, год'


class StyleMixin:
    @property
    def base_style(cls):
        return Titles.base_style

    @property
    def title_style(cls):
        return Titles.title_style

    @property
    def header_style(cls):
        return Titles.header_style

    @property
    def footer_style(cls):
        return Titles.footer_style

    @property
    def header_height(cls):
        return 70

    @property
    def title_row(cls):
        return 1

    @property
    def title_column(cls):
        return 1

    @property
    def header_row(cls):
        return 3

    @property
    def header_column(cls):
        return 1


class WidthsMixin:
    def get_width(self, title):
        widths = {
            Titles.number: 6,
            Titles.tso_name: 16,
            Titles.eto_name: 16,
            Titles.district: 21,
            Titles.source: 16,
            Titles.event_title: 16,
            Titles.terms: 12,
            Titles.laying_type: 16,
            Titles.mw: 7,
            # Titles.gh: 9,
            Titles.th: 7,
            Titles.length: 10,
            Titles.diameter: 7,
            Titles.power: 9,
            Titles.total_cost: 14,
            Titles.total: 10
        }
        return widths[title]


@dataclass
class NetworkEventsBaseTableTitles(WidthsMixin, StyleMixin):
    number: str = Titles.number
    district: str = Titles.district
    source: str = Titles.source
    event_title: str = Titles.event_title
    terms: str = Titles.terms
    # laying_type: str = Titles.laying_type
    length: str = Titles.length
    diametr: str = Titles.diameter
    power: str = Titles.power
    total_cost: str = Titles.total_cost
    total: str = Titles.total

    @property
    def sum_values(cls):
        return {
            Titles.length: 0,
        }


@dataclass
class SourceEventsBaseTableTitles(WidthsMixin, StyleMixin):
    number: str = Titles.number
    source: str = Titles.source
    event_title: str = Titles.event_title
    terms: str = Titles.terms
    mw: str = Titles.mw
    gh: str = Titles.gh
    th: str = Titles.th
    total_cost: str = Titles.total_cost
    total: str = Titles.total

    @property
    def sum_values(cls):
        return {
            Titles.mw: 0,
            Titles.gh: 1,
            Titles.th: 2,
        }
