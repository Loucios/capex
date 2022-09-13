from openpyxl.styles import (NamedStyle,
                             PatternFill,
                             Border,
                             Side,
                             Alignment,
                             Protection,
                             Font)


head_font = Font(
    name='Calibri',
    size=14,
    bold=True,
    italic=False,
    vertAlign=None,
    underline='none',
    strike=False,
    color='FF0000'
)

head_table_font = Font(
    name='Calibri',
    size=11,
    bold=True,
    italic=False,
    vertAlign=None,
    underline='none',
    strike=False,
    color='000000'
)

font = Font(
    name='Calibri',
    size=11,
    bold=False,
    italic=False,
    vertAlign=None,
    underline='none',
    strike=False,
    color='000000'
)

fill = PatternFill(
    fill_type='solid',
    fgColor="ADD8E6",
    # start_color='FFFFFFFF',
    # end_color='FF000000'
)

border = Border(
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

alignment = Alignment(
    horizontal='center',
    vertical='center',
    text_rotation=0,
    wrap_text=True,
    shrink_to_fit=False,
    indent=0
)

number_format = 'General'

protection = Protection(
    locked=True,
    hidden=False
)

base_style = NamedStyle(name='base_style')
base_style.font = font
base_style.border = border
base_style.alignment = alignment

head_table_style = NamedStyle(name='head_table_style')
head_table_style.font = head_table_font
head_table_style.border = border
head_table_style.alignment = alignment
head_table_style.fill = fill

head_style = NamedStyle(name='head_style')
head_style.font = head_font
