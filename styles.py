from dataclasses import dataclass, field

from names import Titles
from openpyxl.styles import (Alignment, Border, Font, NamedStyle, PatternFill,
                             Protection, Side)


@dataclass
class Style:
    title_style: NamedStyle = field(init=False)
    header_style: NamedStyle = field(init=False)
    base_style: NamedStyle = field(init=False)
    footer_style: NamedStyle = field(init=False)

    def __post_init__(self):
        styles = Styles()
        self.title_style = styles.style_1
        self.header_style = styles.style_2
        self.base_style = styles.style_3
        self.footer_style = styles.style_4

    def add_styles(self, wb):
        for self_field in self.__dataclass_fields__:
            wb.add_named_style(getattr(self, self_field))
        return wb


class Styles:
    def __init__(self, titles=None):
        self._titles = Titles() if titles is None else titles

        self._title_font = Font(
            name='Calibri',
            size=14,
            bold=True,
            italic=False,
            vertAlign=None,
            underline='none',
            strike=False,
            color='FF0000'
        )

        self._header_font = Font(
            name='Calibri',
            size=11,
            bold=True,
            italic=False,
            vertAlign=None,
            underline='none',
            strike=False,
            color='000000'
        )

        self._base_font = Font(
            name='Calibri',
            size=11,
            bold=False,
            italic=False,
            vertAlign=None,
            underline='none',
            strike=False,
            color='000000'
        )

        self._footer_font = Font(
            name='Calibri',
            size=11,
            bold=True,
            italic=False,
            vertAlign=None,
            underline='none',
            strike=False,
            color='000000'
        )

        self._header_fill = PatternFill(
            fill_type='solid',
            fgColor="ADD8E6",
            # start_color='FFFFFFFF',
            # end_color='FF000000'
        )

        self._base_border = Border(
            left=Side(border_style='thin', color='000000'),
            right=Side(border_style='thin', color='000000'),
            top=Side(border_style='thin', color='000000'),
            bottom=Side(border_style='thin', color='000000'),
            # diagonal=Side(border_style=None, color='FF000000'),
            # diagonal_direction=0,
            # outline=Side(border_style=None, color='FF000000'),
            # vertical=Side(border_style=None, color='FF000000'),
            # horizontal=Side(border_style=None, color='FF000000')
        )

        self._base_alignment = Alignment(
            horizontal='center',
            vertical='center',
            text_rotation=0,
            wrap_text=True,
            shrink_to_fit=False,
            indent=0
        )

        self._number_format = 'General'

        self._protection = Protection(
            locked=True,
            hidden=False
        )

    @property
    def style_1(self):
        base_style = NamedStyle(name=self._titles.base_style)
        base_style.font = self._base_font
        base_style.border = self._base_border
        base_style.alignment = self._base_alignment
        return base_style

    @property
    def style_2(self):
        footer_style = NamedStyle(name=self._titles.footer_style)
        footer_style.font = self._footer_font
        footer_style.border = self._base_border
        footer_style.alignment = self._base_alignment
        return footer_style

    @property
    def style_3(self):
        header_style = NamedStyle(name=self._titles.header_style)
        header_style.font = self._header_font
        header_style.border = self._base_border
        header_style.alignment = self._base_alignment
        header_style.fill = self._header_fill
        return header_style

    @property
    def style_4(self):
        title_style = NamedStyle(name=self._titles.title_style)
        title_style.font = self._title_font
        return title_style
