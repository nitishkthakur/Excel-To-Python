#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

from excel_pipeline.python_engine_support import (
    execute_unstructured_python_engine,
    xl_add,
    xl_and,
    xl_average,
    xl_concat,
    xl_div,
    xl_eq,
    xl_error,
    xl_ge,
    xl_gt,
    xl_if,
    xl_iferror,
    xl_le,
    xl_lt,
    xl_max,
    xl_min,
    xl_mod,
    xl_mul,
    xl_ne,
    xl_now,
    xl_npv,
    xl_or,
    xl_percent,
    xl_pow,
    xl_ref,
    xl_rounddown,
    xl_roundup,
    xl_sub,
    xl_sum,
    xl_sumproduct,
    xl_today,
    xl_uminus,
    xl_uplus,
    xl_month,
    xl_year,
)

DEFAULT_MAPPING_REPORT = Path(r"/home/nitish/Documents/github/Excel-To-Python/Codex/artifacts_python_engine/Gail-India/mapping_report.xlsx")
DEFAULT_INPUT = Path(r"/home/nitish/Documents/github/Excel-To-Python/Codex/artifacts_python_engine/Gail-India/unstructured_inputs.xlsx")
DEFAULT_OUTPUT = Path(r"/home/nitish/Documents/github/Excel-To-Python/Codex/artifacts_python_engine/Gail-India/unstructured_output_python.xlsx")

def calc_COMPANY_OVERVIEW_G23(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AD41')), xl_ref(ctx.cell('INCOME STATEMENT', 'AD6')))

def calc_COMPANY_OVERVIEW_H23(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AE41')), xl_ref(ctx.cell('INCOME STATEMENT', 'AE6')))

def calc_COMPANY_OVERVIEW_I23(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AF41')), xl_ref(ctx.cell('INCOME STATEMENT', 'AF6')))

def calc_COMPANY_OVERVIEW_J23(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AG41')), xl_ref(ctx.cell('INCOME STATEMENT', 'AG6')))

def calc_COMPANY_OVERVIEW_E28(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'G7')), xl_ref(ctx.cell('INCOME STATEMENT', 'V6')))

def calc_COMPANY_OVERVIEW_G28(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'G7')), xl_ref(ctx.cell('INCOME STATEMENT', 'AD6')))

def calc_COMPANY_OVERVIEW_H28(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'G7')), xl_ref(ctx.cell('INCOME STATEMENT', 'AE6')))

def calc_COMPANY_OVERVIEW_I28(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'G7')), xl_ref(ctx.cell('INCOME STATEMENT', 'AF6')))

def calc_COMPANY_OVERVIEW_J28(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'G7')), xl_ref(ctx.cell('INCOME STATEMENT', 'AG6')))

def calc_COMPANY_OVERVIEW_B39(ctx):
    return xl_sum(ctx.range('COMPANY OVERVIEW', 'B32:B38'))

def calc_COMPANY_OVERVIEW_B57(ctx):
    return xl_sum(ctx.range('COMPANY OVERVIEW', 'B50:B56'))

def calc_COMPANY_OVERVIEW_C57(ctx):
    return xl_sum(ctx.range('COMPANY OVERVIEW', 'C50:C56'))

def calc_COMPANY_OVERVIEW_D57(ctx):
    return xl_sum(ctx.range('COMPANY OVERVIEW', 'D50:D56'))

def calc_COMPANY_OVERVIEW_E57(ctx):
    return xl_sum(ctx.range('COMPANY OVERVIEW', 'E50:E56'))

def calc_COMPANY_OVERVIEW_F57(ctx):
    return xl_sum(ctx.range('COMPANY OVERVIEW', 'F50:F56'))

def calc_COMPANY_OVERVIEW_G57(ctx):
    return xl_sum(ctx.range('COMPANY OVERVIEW', 'G50:G56'))

def calc_PRESENTATION_F8(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'B8:E8'))

def calc_PRESENTATION_K8(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'G8:J8'))

def calc_PRESENTATION_P8(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'L8:O8'))

def calc_PRESENTATION_T8(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'U8')), xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'Q8')), xl_ref(ctx.cell('PRESENTATION', 'R8'))), xl_ref(ctx.cell('PRESENTATION', 'S8'))))

def calc_PRESENTATION_F9(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'B9:E9'))

def calc_PRESENTATION_K9(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'G9:J9'))

def calc_PRESENTATION_P9(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'L9:O9'))

def calc_PRESENTATION_T9(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'U9')), xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'Q9')), xl_ref(ctx.cell('PRESENTATION', 'R9'))), xl_ref(ctx.cell('PRESENTATION', 'S9'))))

def calc_PRESENTATION_F10(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'B10:E10'))

def calc_PRESENTATION_K10(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'G10:J10'))

def calc_PRESENTATION_P10(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'L10:O10'))

def calc_PRESENTATION_T10(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'U10')), xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'Q10')), xl_ref(ctx.cell('PRESENTATION', 'R10'))), xl_ref(ctx.cell('PRESENTATION', 'S10'))))

def calc_PRESENTATION_F11(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'B11:E11'))

def calc_PRESENTATION_K11(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'G11:J11'))

def calc_PRESENTATION_P11(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'L11:O11'))

def calc_PRESENTATION_T11(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'U11')), xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'Q11')), xl_ref(ctx.cell('PRESENTATION', 'R11'))), xl_ref(ctx.cell('PRESENTATION', 'S11'))))

def calc_PRESENTATION_F12(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'B12:E12'))

def calc_PRESENTATION_K12(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'G12:J12'))

def calc_PRESENTATION_P12(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'L12:O12'))

def calc_PRESENTATION_T12(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'U12')), xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'Q12')), xl_ref(ctx.cell('PRESENTATION', 'R12'))), xl_ref(ctx.cell('PRESENTATION', 'S12'))))

def calc_PRESENTATION_F13(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'B13:E13'))

def calc_PRESENTATION_K13(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'G13:J13'))

def calc_PRESENTATION_P13(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'L13:O13'))

def calc_PRESENTATION_T13(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'U13')), xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'Q13')), xl_ref(ctx.cell('PRESENTATION', 'R13'))), xl_ref(ctx.cell('PRESENTATION', 'S13'))))

def calc_PRESENTATION_F14(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'B14:E14'))

def calc_PRESENTATION_K14(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'G14:J14'))

def calc_PRESENTATION_P14(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'L14:O14'))

def calc_PRESENTATION_T14(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'U14')), xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'Q14')), xl_ref(ctx.cell('PRESENTATION', 'R14'))), xl_ref(ctx.cell('PRESENTATION', 'S14'))))

def calc_PRESENTATION_B15(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'B8:B14'))

def calc_PRESENTATION_C15(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'C8:C14'))

def calc_PRESENTATION_D15(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'D8:D14'))

def calc_PRESENTATION_E15(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'E8:E14'))

def calc_PRESENTATION_G15(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'G8:G14'))

def calc_PRESENTATION_H15(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'H8:H14'))

def calc_PRESENTATION_I15(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'I8:I14'))

def calc_PRESENTATION_J15(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'J8:J14'))

def calc_PRESENTATION_L15(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'L8:L14'))

def calc_PRESENTATION_M15(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'M8:M14'))

def calc_PRESENTATION_N15(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'N8:N14'))

def calc_PRESENTATION_O15(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'O8:O14'))

def calc_PRESENTATION_Q15(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'Q8:Q14'))

def calc_PRESENTATION_R15(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'R8:R14'))

def calc_PRESENTATION_S15(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'S8:S14'))

def calc_PRESENTATION_V15(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'V8:V14'))

def calc_PRESENTATION_W15(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'W8:W14'))

def calc_PRESENTATION_X15(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'X8:X14'))

def calc_PRESENTATION_Y15(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'Y8:Y14'))

def calc_PRESENTATION_F16(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'B16:E16'))

def calc_PRESENTATION_P16(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'L16:O16'))

def calc_PRESENTATION_U16(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'Q16:T16'))

def calc_PRESENTATION_D25(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'B25')), xl_ref(ctx.cell('PRESENTATION', 'C25')))

def calc_PRESENTATION_K25(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'I25')), xl_ref(ctx.cell('PRESENTATION', 'J25')))

def calc_PRESENTATION_R25(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'P25')), xl_ref(ctx.cell('PRESENTATION', 'Q25')))

def calc_PRESENTATION_Y25(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'W25')), xl_ref(ctx.cell('PRESENTATION', 'X25')))

def calc_PRESENTATION_B26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'B27')), xl_ref(ctx.cell('PRESENTATION', 'B28'))), xl_ref(ctx.cell('PRESENTATION', 'B29'))), xl_ref(ctx.cell('PRESENTATION', 'B30'))), xl_ref(ctx.cell('PRESENTATION', 'B31')))

def calc_PRESENTATION_C26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'C27')), xl_ref(ctx.cell('PRESENTATION', 'C28'))), xl_ref(ctx.cell('PRESENTATION', 'C29'))), xl_ref(ctx.cell('PRESENTATION', 'C30'))), xl_ref(ctx.cell('PRESENTATION', 'C31')))

def calc_PRESENTATION_E26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'E27')), xl_ref(ctx.cell('PRESENTATION', 'E28'))), xl_ref(ctx.cell('PRESENTATION', 'E29'))), xl_ref(ctx.cell('PRESENTATION', 'E30'))), xl_ref(ctx.cell('PRESENTATION', 'E31')))

def calc_PRESENTATION_G26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'G27')), xl_ref(ctx.cell('PRESENTATION', 'G28'))), xl_ref(ctx.cell('PRESENTATION', 'G29'))), xl_ref(ctx.cell('PRESENTATION', 'G30'))), xl_ref(ctx.cell('PRESENTATION', 'G31')))

def calc_PRESENTATION_I26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'I27')), xl_ref(ctx.cell('PRESENTATION', 'I28'))), xl_ref(ctx.cell('PRESENTATION', 'I30'))), xl_ref(ctx.cell('PRESENTATION', 'I29'))), xl_ref(ctx.cell('PRESENTATION', 'I61')))

def calc_PRESENTATION_J26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'J27')), xl_ref(ctx.cell('PRESENTATION', 'J28'))), xl_ref(ctx.cell('PRESENTATION', 'J30'))), xl_ref(ctx.cell('PRESENTATION', 'J29'))), xl_ref(ctx.cell('PRESENTATION', 'J31')))

def calc_PRESENTATION_L26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'L27')), xl_ref(ctx.cell('PRESENTATION', 'L28'))), xl_ref(ctx.cell('PRESENTATION', 'L30'))), xl_ref(ctx.cell('PRESENTATION', 'L29'))), xl_ref(ctx.cell('PRESENTATION', 'L31')))

def calc_PRESENTATION_N26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'N27')), xl_ref(ctx.cell('PRESENTATION', 'N28'))), xl_ref(ctx.cell('PRESENTATION', 'N30'))), xl_ref(ctx.cell('PRESENTATION', 'N29'))), xl_ref(ctx.cell('PRESENTATION', 'N31')))

def calc_PRESENTATION_P26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'P27')), xl_ref(ctx.cell('PRESENTATION', 'P28'))), xl_ref(ctx.cell('PRESENTATION', 'P30'))), xl_ref(ctx.cell('PRESENTATION', 'P29'))), xl_ref(ctx.cell('PRESENTATION', 'P31')))

def calc_PRESENTATION_Q26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'Q27')), xl_ref(ctx.cell('PRESENTATION', 'Q28'))), xl_ref(ctx.cell('PRESENTATION', 'U30'))), xl_ref(ctx.cell('PRESENTATION', 'Q29'))), xl_ref(ctx.cell('PRESENTATION', 'Q31')))

def calc_PRESENTATION_S26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'S27')), xl_ref(ctx.cell('PRESENTATION', 'S28'))), xl_ref(ctx.cell('PRESENTATION', 'S29'))), xl_ref(ctx.cell('PRESENTATION', 'S30'))), xl_ref(ctx.cell('PRESENTATION', 'S31')))

def calc_PRESENTATION_U26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'U27')), xl_ref(ctx.cell('PRESENTATION', 'U28'))), xl_ref(ctx.cell('PRESENTATION', 'U30'))), xl_ref(ctx.cell('PRESENTATION', 'U29'))), xl_ref(ctx.cell('PRESENTATION', 'U31')))

def calc_PRESENTATION_W26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'W27')), xl_ref(ctx.cell('PRESENTATION', 'W28'))), xl_ref(ctx.cell('PRESENTATION', 'W30'))), xl_ref(ctx.cell('PRESENTATION', 'W29'))), xl_ref(ctx.cell('PRESENTATION', 'W31')))

def calc_PRESENTATION_X26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'X27')), xl_ref(ctx.cell('PRESENTATION', 'X28'))), xl_ref(ctx.cell('PRESENTATION', 'X30'))), xl_ref(ctx.cell('PRESENTATION', 'X29'))), xl_ref(ctx.cell('PRESENTATION', 'X31')))

def calc_PRESENTATION_Z26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'Z27')), xl_ref(ctx.cell('PRESENTATION', 'Z28'))), xl_ref(ctx.cell('PRESENTATION', 'Z30'))), xl_ref(ctx.cell('PRESENTATION', 'Z29'))), xl_ref(ctx.cell('PRESENTATION', 'Z31')))

def calc_PRESENTATION_AB26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'AB27')), xl_ref(ctx.cell('PRESENTATION', 'AB28'))), xl_ref(ctx.cell('PRESENTATION', 'AB30'))), xl_ref(ctx.cell('PRESENTATION', 'AB29'))), xl_ref(ctx.cell('PRESENTATION', 'AB31')))

def calc_PRESENTATION_AD26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'AD27')), xl_ref(ctx.cell('PRESENTATION', 'AD28'))), xl_ref(ctx.cell('PRESENTATION', 'AD29'))), xl_ref(ctx.cell('PRESENTATION', 'AD30'))), xl_ref(ctx.cell('PRESENTATION', 'AD31')))

def calc_PRESENTATION_AE26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'AE27')), xl_ref(ctx.cell('PRESENTATION', 'AE28'))), xl_ref(ctx.cell('PRESENTATION', 'AE29'))), xl_ref(ctx.cell('PRESENTATION', 'AE30'))), xl_ref(ctx.cell('PRESENTATION', 'AE31')))

def calc_PRESENTATION_AF26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'AF27')), xl_ref(ctx.cell('PRESENTATION', 'AF28'))), xl_ref(ctx.cell('PRESENTATION', 'AF29'))), xl_ref(ctx.cell('PRESENTATION', 'AF30'))), xl_ref(ctx.cell('PRESENTATION', 'AF31')))

def calc_PRESENTATION_AG26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'AG27')), xl_ref(ctx.cell('PRESENTATION', 'AG28'))), xl_ref(ctx.cell('PRESENTATION', 'AG29'))), xl_ref(ctx.cell('PRESENTATION', 'AG30'))), xl_ref(ctx.cell('PRESENTATION', 'AG31')))

def calc_PRESENTATION_D27(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'B27')), xl_ref(ctx.cell('PRESENTATION', 'C27')))

def calc_PRESENTATION_K27(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'I27')), xl_ref(ctx.cell('PRESENTATION', 'J27')))

def calc_PRESENTATION_R27(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'P27')), xl_ref(ctx.cell('PRESENTATION', 'Q27')))

def calc_PRESENTATION_Y27(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'W27')), xl_ref(ctx.cell('PRESENTATION', 'X27')))

def calc_PRESENTATION_D28(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'B28')), xl_ref(ctx.cell('PRESENTATION', 'C28')))

def calc_PRESENTATION_K28(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'I28')), xl_ref(ctx.cell('PRESENTATION', 'J28')))

def calc_PRESENTATION_R28(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'P28')), xl_ref(ctx.cell('PRESENTATION', 'Q28')))

def calc_PRESENTATION_Y28(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'W28')), xl_ref(ctx.cell('PRESENTATION', 'X28')))

def calc_PRESENTATION_D29(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'B29')), xl_ref(ctx.cell('PRESENTATION', 'C29')))

def calc_PRESENTATION_K29(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'I29')), xl_ref(ctx.cell('PRESENTATION', 'J29')))

def calc_PRESENTATION_R29(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'P29')), xl_ref(ctx.cell('PRESENTATION', 'Q29')))

def calc_PRESENTATION_Y29(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'W29')), xl_ref(ctx.cell('PRESENTATION', 'X29')))

def calc_PRESENTATION_D30(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'B30')), xl_ref(ctx.cell('PRESENTATION', 'C30')))

def calc_PRESENTATION_K30(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'I30')), xl_ref(ctx.cell('PRESENTATION', 'J30')))

def calc_PRESENTATION_R30(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'P30')), xl_ref(ctx.cell('PRESENTATION', 'U30')))

def calc_PRESENTATION_Y30(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'W30')), xl_ref(ctx.cell('PRESENTATION', 'X30')))

def calc_PRESENTATION_D31(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'B31')), xl_ref(ctx.cell('PRESENTATION', 'C31')))

def calc_PRESENTATION_K31(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'I61')), xl_ref(ctx.cell('PRESENTATION', 'J31')))

def calc_PRESENTATION_R31(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'P31')), xl_ref(ctx.cell('PRESENTATION', 'Q31')))

def calc_PRESENTATION_Y31(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'W31')), xl_ref(ctx.cell('PRESENTATION', 'X31')))

def calc_PRESENTATION_K32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'L25')), xl_ref(ctx.cell('PRESENTATION', 'K26')))

def calc_PRESENTATION_D34(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'B34')), xl_ref(ctx.cell('PRESENTATION', 'C34')))

def calc_PRESENTATION_K34(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'I34')), xl_ref(ctx.cell('PRESENTATION', 'J34')))

def calc_PRESENTATION_R34(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'P34')), xl_ref(ctx.cell('PRESENTATION', 'Q34')))

def calc_PRESENTATION_Y34(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'W34')), xl_ref(ctx.cell('PRESENTATION', 'X34')))

def calc_PRESENTATION_D35(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'B35')), xl_ref(ctx.cell('PRESENTATION', 'C35')))

def calc_PRESENTATION_K35(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'I35')), xl_ref(ctx.cell('PRESENTATION', 'J35')))

def calc_PRESENTATION_R35(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'P35')), xl_ref(ctx.cell('PRESENTATION', 'Q35')))

def calc_PRESENTATION_Y35(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'W35')), xl_ref(ctx.cell('PRESENTATION', 'X35')))

def calc_PRESENTATION_D37(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'B37')), xl_ref(ctx.cell('PRESENTATION', 'C37')))

def calc_PRESENTATION_K37(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'I37')), xl_ref(ctx.cell('PRESENTATION', 'J37')))

def calc_PRESENTATION_R37(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'P37')), xl_ref(ctx.cell('PRESENTATION', 'Q37')))

def calc_PRESENTATION_V37(ctx):
    return xl_add(61177, 18485)

def calc_PRESENTATION_Y37(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'W37')), xl_ref(ctx.cell('PRESENTATION', 'X37')))

def calc_PRESENTATION_M39(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'K39')), xl_ref(ctx.cell('PRESENTATION', 'L39')))

def calc_PRESENTATION_R39(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'P39')), xl_ref(ctx.cell('PRESENTATION', 'Q39')))

def calc_PRESENTATION_Y39(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'W39')), xl_ref(ctx.cell('PRESENTATION', 'X39')))

def calc_PRESENTATION_D41(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'B41')), xl_ref(ctx.cell('PRESENTATION', 'C41')))

def calc_PRESENTATION_K41(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'I41')), xl_ref(ctx.cell('PRESENTATION', 'J41')))

def calc_PRESENTATION_R41(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'P41')), xl_ref(ctx.cell('PRESENTATION', 'Q41')))

def calc_PRESENTATION_Y41(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'W41')), xl_ref(ctx.cell('PRESENTATION', 'X41')))

def calc_PRESENTATION_D42(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'B42')), xl_ref(ctx.cell('PRESENTATION', 'C42')))

def calc_PRESENTATION_K42(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'I42')), xl_ref(ctx.cell('PRESENTATION', 'J42')))

def calc_PRESENTATION_R42(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'P42')), xl_ref(ctx.cell('PRESENTATION', 'Q42')))

def calc_PRESENTATION_Y42(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'W42')), xl_ref(ctx.cell('PRESENTATION', 'X42')))

def calc_PRESENTATION_D43(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'B43')), xl_ref(ctx.cell('PRESENTATION', 'C43')))

def calc_PRESENTATION_K43(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'I43')), xl_ref(ctx.cell('PRESENTATION', 'J43')))

def calc_PRESENTATION_R43(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'P43')), xl_ref(ctx.cell('PRESENTATION', 'Q43')))

def calc_PRESENTATION_Y43(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'W43')), xl_ref(ctx.cell('PRESENTATION', 'X43')))

def calc_Segment_Revenue_Model_F8(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'B8:E8'))

def calc_Segment_Revenue_Model_K8(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'G8:J8'))

def calc_Segment_Revenue_Model_P8(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L8:O8'))

def calc_Segment_Revenue_Model_F9(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'B9:E9'))

def calc_Segment_Revenue_Model_K9(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'G9:J9'))

def calc_Segment_Revenue_Model_P9(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L9:O9'))

def calc_Segment_Revenue_Model_F10(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'B10:E10'))

def calc_Segment_Revenue_Model_K10(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'G10:J10'))

def calc_Segment_Revenue_Model_P10(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L10:O10'))

def calc_Segment_Revenue_Model_F11(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'B11:E11'))

def calc_Segment_Revenue_Model_K11(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'G11:J11'))

def calc_Segment_Revenue_Model_P11(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L11:O11'))

def calc_Segment_Revenue_Model_F12(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'B12:E12'))

def calc_Segment_Revenue_Model_K12(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'G12:J12'))

def calc_Segment_Revenue_Model_P12(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L12:O12'))

def calc_Segment_Revenue_Model_U12(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'U28')), 1), 3819)

def calc_Segment_Revenue_Model_F13(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'B13:E13'))

def calc_Segment_Revenue_Model_K13(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'G13:J13'))

def calc_Segment_Revenue_Model_P13(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L13:O13'))

def calc_Segment_Revenue_Model_F14(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'B14:E14'))

def calc_Segment_Revenue_Model_K14(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'G14:J14'))

def calc_Segment_Revenue_Model_P14(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L14:O14'))

def calc_Segment_Revenue_Model_B15(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'B8:B14'))

def calc_Segment_Revenue_Model_C15(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'C8:C14'))

def calc_Segment_Revenue_Model_D15(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'D8:D14'))

def calc_Segment_Revenue_Model_E15(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'E8:E14'))

def calc_Segment_Revenue_Model_G15(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'G8:G14'))

def calc_Segment_Revenue_Model_H15(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'H8:H14'))

def calc_Segment_Revenue_Model_I15(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'I8:I14'))

def calc_Segment_Revenue_Model_J15(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'J8:J14'))

def calc_Segment_Revenue_Model_L15(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L8:L14'))

def calc_Segment_Revenue_Model_M15(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'M8:M14'))

def calc_Segment_Revenue_Model_N15(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'N8:N14'))

def calc_Segment_Revenue_Model_O15(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'O8:O14'))

def calc_Segment_Revenue_Model_Q15(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'Q8:Q14'))

def calc_Segment_Revenue_Model_R15(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'R8:R14'))

def calc_Segment_Revenue_Model_S15(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'S8:S14'))

def calc_Segment_Revenue_Model_F16(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'B16:E16'))

def calc_Segment_Revenue_Model_P16(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L16:O16'))

def calc_Segment_Revenue_Model_U16(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'Q16:T16'))

def calc_Segment_Revenue_Model_G24(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'G8')), xl_ref(ctx.cell('Segment Revenue Model', 'B8'))), 1)

def calc_Segment_Revenue_Model_H24(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'H8')), xl_ref(ctx.cell('Segment Revenue Model', 'C8'))), 1)

def calc_Segment_Revenue_Model_I24(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'I8')), xl_ref(ctx.cell('Segment Revenue Model', 'D8'))), 1)

def calc_Segment_Revenue_Model_J24(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'J8')), xl_ref(ctx.cell('Segment Revenue Model', 'E8'))), 1)

def calc_Segment_Revenue_Model_L24(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'L8')), xl_ref(ctx.cell('Segment Revenue Model', 'G8'))), 1)

def calc_Segment_Revenue_Model_M24(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'M8')), xl_ref(ctx.cell('Segment Revenue Model', 'H8'))), 1)

def calc_Segment_Revenue_Model_N24(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'N8')), xl_ref(ctx.cell('Segment Revenue Model', 'I8'))), 1)

def calc_Segment_Revenue_Model_O24(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'O8')), xl_ref(ctx.cell('Segment Revenue Model', 'J8'))), 1)

def calc_Segment_Revenue_Model_Q24(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Q8')), xl_ref(ctx.cell('Segment Revenue Model', 'L8'))), 1)

def calc_Segment_Revenue_Model_R24(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'R8')), xl_ref(ctx.cell('Segment Revenue Model', 'M8'))), 1)

def calc_Segment_Revenue_Model_S24(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'S8')), xl_ref(ctx.cell('Segment Revenue Model', 'N8'))), 1)

def calc_Segment_Revenue_Model_G25(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'G9')), xl_ref(ctx.cell('Segment Revenue Model', 'B9'))), 1)

def calc_Segment_Revenue_Model_H25(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'H9')), xl_ref(ctx.cell('Segment Revenue Model', 'C9'))), 1)

def calc_Segment_Revenue_Model_I25(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'I9')), xl_ref(ctx.cell('Segment Revenue Model', 'D9'))), 1)

def calc_Segment_Revenue_Model_J25(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'J9')), xl_ref(ctx.cell('Segment Revenue Model', 'E9'))), 1)

def calc_Segment_Revenue_Model_L25(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'L9')), xl_ref(ctx.cell('Segment Revenue Model', 'G9'))), 1)

def calc_Segment_Revenue_Model_M25(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'M9')), xl_ref(ctx.cell('Segment Revenue Model', 'H9'))), 1)

def calc_Segment_Revenue_Model_N25(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'N9')), xl_ref(ctx.cell('Segment Revenue Model', 'I9'))), 1)

def calc_Segment_Revenue_Model_O25(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'O9')), xl_ref(ctx.cell('Segment Revenue Model', 'J9'))), 1)

def calc_Segment_Revenue_Model_Q25(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Q9')), xl_ref(ctx.cell('Segment Revenue Model', 'L9'))), 1)

def calc_Segment_Revenue_Model_R25(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'R9')), xl_ref(ctx.cell('Segment Revenue Model', 'M9'))), 1)

def calc_Segment_Revenue_Model_S25(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'S9')), xl_ref(ctx.cell('Segment Revenue Model', 'N9'))), 1)

def calc_Segment_Revenue_Model_G26(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'G10')), xl_ref(ctx.cell('Segment Revenue Model', 'B10'))), 1)

def calc_Segment_Revenue_Model_H26(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'H10')), xl_ref(ctx.cell('Segment Revenue Model', 'C10'))), 1)

def calc_Segment_Revenue_Model_I26(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'I10')), xl_ref(ctx.cell('Segment Revenue Model', 'D10'))), 1)

def calc_Segment_Revenue_Model_J26(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'J10')), xl_ref(ctx.cell('Segment Revenue Model', 'E10'))), 1)

def calc_Segment_Revenue_Model_L26(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'L10')), xl_ref(ctx.cell('Segment Revenue Model', 'G10'))), 1)

def calc_Segment_Revenue_Model_M26(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'M10')), xl_ref(ctx.cell('Segment Revenue Model', 'H10'))), 1)

def calc_Segment_Revenue_Model_N26(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'N10')), xl_ref(ctx.cell('Segment Revenue Model', 'I10'))), 1)

def calc_Segment_Revenue_Model_O26(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'O10')), xl_ref(ctx.cell('Segment Revenue Model', 'J10'))), 1)

def calc_Segment_Revenue_Model_Q26(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Q10')), xl_ref(ctx.cell('Segment Revenue Model', 'L10'))), 1)

def calc_Segment_Revenue_Model_R26(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'R10')), xl_ref(ctx.cell('Segment Revenue Model', 'M10'))), 1)

def calc_Segment_Revenue_Model_S26(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'S10')), xl_ref(ctx.cell('Segment Revenue Model', 'N10'))), 1)

def calc_Segment_Revenue_Model_G27(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'G11')), xl_ref(ctx.cell('Segment Revenue Model', 'B11'))), 1)

def calc_Segment_Revenue_Model_H27(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'H11')), xl_ref(ctx.cell('Segment Revenue Model', 'C11'))), 1)

def calc_Segment_Revenue_Model_I27(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'I11')), xl_ref(ctx.cell('Segment Revenue Model', 'D11'))), 1)

def calc_Segment_Revenue_Model_J27(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'J11')), xl_ref(ctx.cell('Segment Revenue Model', 'E11'))), 1)

def calc_Segment_Revenue_Model_L27(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'L11')), xl_ref(ctx.cell('Segment Revenue Model', 'G11'))), 1)

def calc_Segment_Revenue_Model_M27(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'M11')), xl_ref(ctx.cell('Segment Revenue Model', 'H11'))), 1)

def calc_Segment_Revenue_Model_N27(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'N11')), xl_ref(ctx.cell('Segment Revenue Model', 'I11'))), 1)

def calc_Segment_Revenue_Model_O27(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'O11')), xl_ref(ctx.cell('Segment Revenue Model', 'J11'))), 1)

def calc_Segment_Revenue_Model_Q27(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Q11')), xl_ref(ctx.cell('Segment Revenue Model', 'L11'))), 1)

def calc_Segment_Revenue_Model_R27(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'R11')), xl_ref(ctx.cell('Segment Revenue Model', 'M11'))), 1)

def calc_Segment_Revenue_Model_S27(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'S11')), xl_ref(ctx.cell('Segment Revenue Model', 'N11'))), 1)

def calc_Segment_Revenue_Model_G28(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'G12')), xl_ref(ctx.cell('Segment Revenue Model', 'B12'))), 1)

def calc_Segment_Revenue_Model_H28(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'H12')), xl_ref(ctx.cell('Segment Revenue Model', 'C12'))), 1)

def calc_Segment_Revenue_Model_I28(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'I12')), xl_ref(ctx.cell('Segment Revenue Model', 'D12'))), 1)

def calc_Segment_Revenue_Model_J28(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'J12')), xl_ref(ctx.cell('Segment Revenue Model', 'E12'))), 1)

def calc_Segment_Revenue_Model_L28(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'L12')), xl_ref(ctx.cell('Segment Revenue Model', 'G12'))), 1)

def calc_Segment_Revenue_Model_M28(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'M12')), xl_ref(ctx.cell('Segment Revenue Model', 'H12'))), 1)

def calc_Segment_Revenue_Model_N28(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'N12')), xl_ref(ctx.cell('Segment Revenue Model', 'I12'))), 1)

def calc_Segment_Revenue_Model_O28(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'O12')), xl_ref(ctx.cell('Segment Revenue Model', 'J12'))), 1)

def calc_Segment_Revenue_Model_Q28(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Q12')), xl_ref(ctx.cell('Segment Revenue Model', 'L12'))), 1)

def calc_Segment_Revenue_Model_R28(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'R12')), xl_ref(ctx.cell('Segment Revenue Model', 'M12'))), 1)

def calc_Segment_Revenue_Model_S28(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'S12')), xl_ref(ctx.cell('Segment Revenue Model', 'N12'))), 1)

def calc_Segment_Revenue_Model_G29(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'G13')), xl_ref(ctx.cell('Segment Revenue Model', 'B13'))), 1)

def calc_Segment_Revenue_Model_H29(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'H13')), xl_ref(ctx.cell('Segment Revenue Model', 'C13'))), 1)

def calc_Segment_Revenue_Model_I29(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'I13')), xl_ref(ctx.cell('Segment Revenue Model', 'D13'))), 1)

def calc_Segment_Revenue_Model_J29(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'J13')), xl_ref(ctx.cell('Segment Revenue Model', 'E13'))), 1)

def calc_Segment_Revenue_Model_L29(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'L13')), xl_ref(ctx.cell('Segment Revenue Model', 'G13'))), 1)

def calc_Segment_Revenue_Model_M29(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'M13')), xl_ref(ctx.cell('Segment Revenue Model', 'H13'))), 1)

def calc_Segment_Revenue_Model_N29(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'N13')), xl_ref(ctx.cell('Segment Revenue Model', 'I13'))), 1)

def calc_Segment_Revenue_Model_O29(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'O13')), xl_ref(ctx.cell('Segment Revenue Model', 'J13'))), 1)

def calc_Segment_Revenue_Model_Q29(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Q13')), xl_ref(ctx.cell('Segment Revenue Model', 'L13'))), 1)

def calc_Segment_Revenue_Model_R29(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'R13')), xl_ref(ctx.cell('Segment Revenue Model', 'M13'))), 1)

def calc_Segment_Revenue_Model_S29(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'S13')), xl_ref(ctx.cell('Segment Revenue Model', 'N13'))), 1)

def calc_Segment_Revenue_Model_G30(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'G14')), xl_ref(ctx.cell('Segment Revenue Model', 'B14'))), 1)

def calc_Segment_Revenue_Model_H30(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'H14')), xl_ref(ctx.cell('Segment Revenue Model', 'C14'))), 1)

def calc_Segment_Revenue_Model_I30(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'I14')), xl_ref(ctx.cell('Segment Revenue Model', 'D14'))), 1)

def calc_Segment_Revenue_Model_J30(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'J14')), xl_ref(ctx.cell('Segment Revenue Model', 'E14'))), 1)

def calc_Segment_Revenue_Model_L30(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'L14')), xl_ref(ctx.cell('Segment Revenue Model', 'G14'))), 1)

def calc_Segment_Revenue_Model_M30(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'M14')), xl_ref(ctx.cell('Segment Revenue Model', 'H14'))), 1)

def calc_Segment_Revenue_Model_N30(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'N14')), xl_ref(ctx.cell('Segment Revenue Model', 'I14'))), 1)

def calc_Segment_Revenue_Model_O30(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'O14')), xl_ref(ctx.cell('Segment Revenue Model', 'J14'))), 1)

def calc_Segment_Revenue_Model_Q30(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Q14')), xl_ref(ctx.cell('Segment Revenue Model', 'L14'))), 1)

def calc_Segment_Revenue_Model_R30(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'R14')), xl_ref(ctx.cell('Segment Revenue Model', 'M14'))), 1)

def calc_Segment_Revenue_Model_S30(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'S14')), xl_ref(ctx.cell('Segment Revenue Model', 'N14'))), 1)

def calc_Segment_Revenue_Model_G32(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'G16')), xl_ref(ctx.cell('Segment Revenue Model', 'B16'))), 1)

def calc_Segment_Revenue_Model_H32(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'H16')), xl_ref(ctx.cell('Segment Revenue Model', 'C16'))), 1)

def calc_Segment_Revenue_Model_I32(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'I16')), xl_ref(ctx.cell('Segment Revenue Model', 'D16'))), 1)

def calc_Segment_Revenue_Model_J32(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'J16')), xl_ref(ctx.cell('Segment Revenue Model', 'E16'))), 1)

def calc_Segment_Revenue_Model_L32(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'L16')), xl_ref(ctx.cell('Segment Revenue Model', 'G16'))), 1)

def calc_Segment_Revenue_Model_M32(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'M16')), xl_ref(ctx.cell('Segment Revenue Model', 'H16'))), 1)

def calc_Segment_Revenue_Model_N32(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'N16')), xl_ref(ctx.cell('Segment Revenue Model', 'I16'))), 1)

def calc_Segment_Revenue_Model_O32(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'O16')), xl_ref(ctx.cell('Segment Revenue Model', 'J16'))), 1)

def calc_Segment_Revenue_Model_Q32(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Q16')), xl_ref(ctx.cell('Segment Revenue Model', 'L16'))), 1)

def calc_Segment_Revenue_Model_R32(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'R16')), xl_ref(ctx.cell('Segment Revenue Model', 'M16'))), 1)

def calc_Segment_Revenue_Model_S32(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'S16')), xl_ref(ctx.cell('Segment Revenue Model', 'N16'))), 1)

def calc_Segment_Revenue_Model_T32(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'T16')), xl_ref(ctx.cell('Segment Revenue Model', 'O16'))), 1)

def calc_Segment_Revenue_Model_F51(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'C51:E51'))

def calc_Segment_Revenue_Model_K51(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'G51')), xl_ref(ctx.cell('Segment Revenue Model', 'H51'))), xl_ref(ctx.cell('Segment Revenue Model', 'I51'))), xl_ref(ctx.cell('Segment Revenue Model', 'I52'))), xl_ref(ctx.cell('Segment Revenue Model', 'J52'))), xl_ref(ctx.cell('Segment Revenue Model', 'J51')))

def calc_Segment_Revenue_Model_P51(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L51:O51'))

def calc_Segment_Revenue_Model_U51(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'Q51:T51'))

def calc_Segment_Revenue_Model_F52(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'C52:E52'))

def calc_Segment_Revenue_Model_P52(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L52:O52'))

def calc_Segment_Revenue_Model_U52(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'Q52:T52'))

def calc_Segment_Revenue_Model_F53(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'C53:E53'))

def calc_Segment_Revenue_Model_K53(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'G53:J53'))

def calc_Segment_Revenue_Model_P53(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L53:O53'))

def calc_Segment_Revenue_Model_U53(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'Q53:T53'))

def calc_Segment_Revenue_Model_F54(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'C54:E54'))

def calc_Segment_Revenue_Model_K54(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'G54:J54'))

def calc_Segment_Revenue_Model_P54(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L54:O54'))

def calc_Segment_Revenue_Model_U54(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'Q54:T54'))

def calc_Segment_Revenue_Model_F55(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'C55:E55'))

def calc_Segment_Revenue_Model_K55(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'G55:J55'))

def calc_Segment_Revenue_Model_P55(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L55:O55'))

def calc_Segment_Revenue_Model_U55(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'Q55:T55'))

def calc_Segment_Revenue_Model_F56(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'C56:E56'))

def calc_Segment_Revenue_Model_K56(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'G56:J56'))

def calc_Segment_Revenue_Model_P56(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L56:O56'))

def calc_Segment_Revenue_Model_U56(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'Q56:T56'))

def calc_Segment_Revenue_Model_F57(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'B57:E57'))

def calc_Segment_Revenue_Model_K57(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'G57:J57'))

def calc_Segment_Revenue_Model_P57(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L57:O57'))

def calc_Segment_Revenue_Model_U57(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'Q57:T57'))

def calc_Segment_Revenue_Model_B58(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'B51:B57'))

def calc_Segment_Revenue_Model_C58(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'C51:C57'))

def calc_Segment_Revenue_Model_D58(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'D51:D57'))

def calc_Segment_Revenue_Model_E58(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'E51:E57'))

def calc_Segment_Revenue_Model_G58(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'G53:G57'))

def calc_Segment_Revenue_Model_H58(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'H51:H57'))

def calc_Segment_Revenue_Model_I58(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'I51:I57'))

def calc_Segment_Revenue_Model_J58(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'J51:J57'))

def calc_Segment_Revenue_Model_L58(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L51:L57'))

def calc_Segment_Revenue_Model_M58(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'M51:M57'))

def calc_Segment_Revenue_Model_N58(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'N51:N57'))

def calc_Segment_Revenue_Model_O58(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'O51:O57'))

def calc_Segment_Revenue_Model_Q58(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'Q51:Q57'))

def calc_Segment_Revenue_Model_R58(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'R51:R57'))

def calc_Segment_Revenue_Model_S58(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'S51:S57'))

def calc_Segment_Revenue_Model_T58(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'T51:T57'))

def calc_Segment_Revenue_Model_V58(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'V51:V57'))

def calc_Segment_Revenue_Model_W58(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'W51:W57'))

def calc_Segment_Revenue_Model_X58(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'X51:X57'))

def calc_Segment_Revenue_Model_Y58(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'Y51:Y57'))

def calc_Segment_Revenue_Model_F59(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'B59:E59'))

def calc_Segment_Revenue_Model_K59(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'G59:J59'))

def calc_Segment_Revenue_Model_P59(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L59:O59'))

def calc_Segment_Revenue_Model_U59(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'Q59:T59'))

def calc_Segment_Revenue_Model_F60(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'B60:E60'))

def calc_Segment_Revenue_Model_K60(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'G60:J60'))

def calc_Segment_Revenue_Model_P60(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L60:O60'))

def calc_Segment_Revenue_Model_U60(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'Q60:T60'))

def calc_Segment_Revenue_Model_L61(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L59:L60'))

def calc_Segment_Revenue_Model_F65(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'B65:E65'))

def calc_Segment_Revenue_Model_K65(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'G65:J65'))

def calc_Segment_Revenue_Model_P65(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L65:O65'))

def calc_Segment_Revenue_Model_U65(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'Q65:T65'))

def calc_Segment_Revenue_Model_F66(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'B66:E66'))

def calc_Segment_Revenue_Model_K66(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'G66:J66'))

def calc_Segment_Revenue_Model_P66(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L66:O66'))

def calc_Segment_Revenue_Model_U66(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'Q66:T66'))

def calc_Segment_Revenue_Model_F67(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'B67:E67'))

def calc_Segment_Revenue_Model_K67(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'G67:J67'))

def calc_Segment_Revenue_Model_P67(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L67:O67'))

def calc_Segment_Revenue_Model_U67(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'Q67:T67'))

def calc_Segment_Revenue_Model_F68(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'B68:E68'))

def calc_Segment_Revenue_Model_K68(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'G68:J68'))

def calc_Segment_Revenue_Model_P68(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L68:O68'))

def calc_Segment_Revenue_Model_U68(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'Q68:T68'))

def calc_Segment_Revenue_Model_F69(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'B69:E69'))

def calc_Segment_Revenue_Model_K69(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'G69:J69'))

def calc_Segment_Revenue_Model_P69(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L69:O69'))

def calc_Segment_Revenue_Model_U69(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'Q69:T69'))

def calc_Segment_Revenue_Model_F70(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'B70:E70'))

def calc_Segment_Revenue_Model_K70(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'G70:J70'))

def calc_Segment_Revenue_Model_P70(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L70:O70'))

def calc_Segment_Revenue_Model_U70(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'Q70:T70'))

def calc_Segment_Revenue_Model_F71(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'B71:E71'))

def calc_Segment_Revenue_Model_K71(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'G71:J71'))

def calc_Segment_Revenue_Model_P71(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L71:O71'))

def calc_Segment_Revenue_Model_U71(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'Q71:T71'))

def calc_Segment_Revenue_Model_B72(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'B65:B71'))

def calc_Segment_Revenue_Model_C72(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'C66:C71'))

def calc_Segment_Revenue_Model_D72(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'D65:D71'))

def calc_Segment_Revenue_Model_E72(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'E65:E71'))

def calc_Segment_Revenue_Model_G72(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'G65:G71'))

def calc_Segment_Revenue_Model_H72(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'H65:H71'))

def calc_Segment_Revenue_Model_I72(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'I65:I71'))

def calc_Segment_Revenue_Model_J72(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'J65:J71'))

def calc_Segment_Revenue_Model_L72(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L66:L71'))

def calc_Segment_Revenue_Model_M72(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'M66:M71'))

def calc_Segment_Revenue_Model_N72(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'N65:N71'))

def calc_Segment_Revenue_Model_O72(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'O65:O71'))

def calc_Segment_Revenue_Model_Q72(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'Q65:Q71'))

def calc_Segment_Revenue_Model_R72(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'R65:R71'))

def calc_Segment_Revenue_Model_S72(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'S65:S71'))

def calc_Segment_Revenue_Model_T72(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'T65:T71'))

def calc_Segment_Revenue_Model_F73(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'B73:E73'))

def calc_Segment_Revenue_Model_K73(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'G73:J73'))

def calc_Segment_Revenue_Model_P73(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L73:O73'))

def calc_Segment_Revenue_Model_U73(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'Q73:T73'))

def calc_INCOME_STATEMENT_D6(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'B6')), xl_ref(ctx.cell('INCOME STATEMENT', 'C6')))

def calc_INCOME_STATEMENT_K6(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'I6')), xl_ref(ctx.cell('INCOME STATEMENT', 'J6')))

def calc_INCOME_STATEMENT_R6(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'P6')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q6')))

def calc_INCOME_STATEMENT_Y6(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'W6')), xl_ref(ctx.cell('INCOME STATEMENT', 'X6')))

def calc_INCOME_STATEMENT_B7(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'B6')), xl_ref(ctx.cell('INCOME STATEMENT', 'H6'))), 100)

def calc_INCOME_STATEMENT_C7(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'C6')), xl_ref(ctx.cell('INCOME STATEMENT', 'H6'))), 100)

def calc_INCOME_STATEMENT_E7(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'E6')), xl_ref(ctx.cell('INCOME STATEMENT', 'H6'))), 100)

def calc_INCOME_STATEMENT_G7(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'G6')), xl_ref(ctx.cell('INCOME STATEMENT', 'H6'))), 100)

def calc_INCOME_STATEMENT_P7(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'P6')), xl_ref(ctx.cell('INCOME STATEMENT', 'V6'))), 100)

def calc_INCOME_STATEMENT_Q7(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Q6')), xl_ref(ctx.cell('INCOME STATEMENT', 'V6'))), 100)

def calc_INCOME_STATEMENT_S7(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'S6')), xl_ref(ctx.cell('INCOME STATEMENT', 'V6'))), 100)

def calc_INCOME_STATEMENT_U7(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'U6')), xl_ref(ctx.cell('INCOME STATEMENT', 'V6'))), 100)

def calc_INCOME_STATEMENT_H8(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'H6')), xl_error('#REF!')), 1), 100)

def calc_INCOME_STATEMENT_I8(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'I6')), xl_ref(ctx.cell('INCOME STATEMENT', 'B6'))), 1), 100)

def calc_INCOME_STATEMENT_J8(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'J6')), xl_ref(ctx.cell('INCOME STATEMENT', 'C6'))), 1), 100)

def calc_INCOME_STATEMENT_L8(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'L6')), xl_ref(ctx.cell('INCOME STATEMENT', 'E6'))), xl_mul(1, 100))

def calc_INCOME_STATEMENT_N8(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'N6')), xl_ref(ctx.cell('INCOME STATEMENT', 'G6'))), 1), 100)

def calc_INCOME_STATEMENT_P8(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'P6')), xl_ref(ctx.cell('INCOME STATEMENT', 'I6'))), 1), 100)

def calc_INCOME_STATEMENT_Q8(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Q6')), xl_ref(ctx.cell('INCOME STATEMENT', 'J6'))), 1), 100)

def calc_INCOME_STATEMENT_S8(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'S6')), xl_ref(ctx.cell('INCOME STATEMENT', 'L6'))), 1), 100)

def calc_INCOME_STATEMENT_U8(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'U6')), xl_ref(ctx.cell('INCOME STATEMENT', 'N6'))), 1), 100)

def calc_INCOME_STATEMENT_W8(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'W6')), xl_ref(ctx.cell('INCOME STATEMENT', 'P6'))), 1), 100)

def calc_INCOME_STATEMENT_X8(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'X6')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q6'))), 1), 100)

def calc_INCOME_STATEMENT_Z8(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Z6')), xl_ref(ctx.cell('INCOME STATEMENT', 'S6'))), 1), 100)

def calc_INCOME_STATEMENT_AB8(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AB6')), xl_ref(ctx.cell('INCOME STATEMENT', 'U6'))), 1), 100)

def calc_INCOME_STATEMENT_AE8(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AE6')), xl_ref(ctx.cell('INCOME STATEMENT', 'AD6'))), 1), 100)

def calc_INCOME_STATEMENT_AF8(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AF6')), xl_ref(ctx.cell('INCOME STATEMENT', 'AE6'))), 1), 100)

def calc_INCOME_STATEMENT_AG8(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AG6')), xl_ref(ctx.cell('INCOME STATEMENT', 'AF6'))), 1), 100)

def calc_INCOME_STATEMENT_C9(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'C6')), xl_ref(ctx.cell('INCOME STATEMENT', 'B6'))), 1), 100)

def calc_INCOME_STATEMENT_E9(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'E6')), xl_ref(ctx.cell('INCOME STATEMENT', 'C6'))), 1), 100)

def calc_INCOME_STATEMENT_G9(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'G6')), xl_ref(ctx.cell('INCOME STATEMENT', 'E6'))), 1), 100)

def calc_INCOME_STATEMENT_I9(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'I6')), xl_ref(ctx.cell('INCOME STATEMENT', 'G6'))), 1), 100)

def calc_INCOME_STATEMENT_J9(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'J6')), xl_ref(ctx.cell('INCOME STATEMENT', 'I6'))), 1), 100)

def calc_INCOME_STATEMENT_L9(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'L6')), xl_ref(ctx.cell('INCOME STATEMENT', 'J6'))), 1), 100)

def calc_INCOME_STATEMENT_N9(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'N6')), xl_ref(ctx.cell('INCOME STATEMENT', 'L6'))), 1), 100)

def calc_INCOME_STATEMENT_P9(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'P6')), xl_ref(ctx.cell('INCOME STATEMENT', 'N6'))), 1), 100)

def calc_INCOME_STATEMENT_Q9(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Q6')), xl_ref(ctx.cell('INCOME STATEMENT', 'P6'))), 1), 100)

def calc_INCOME_STATEMENT_S9(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'S6')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q6'))), 1), 100)

def calc_INCOME_STATEMENT_U9(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'U6')), xl_ref(ctx.cell('INCOME STATEMENT', 'S6'))), 1), 100)

def calc_INCOME_STATEMENT_W9(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'W6')), xl_ref(ctx.cell('INCOME STATEMENT', 'U6'))), 1), 100)

def calc_INCOME_STATEMENT_X9(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'X6')), xl_ref(ctx.cell('INCOME STATEMENT', 'W6'))), 1), 100)

def calc_INCOME_STATEMENT_Z9(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Z6')), xl_ref(ctx.cell('INCOME STATEMENT', 'X6'))), 1), 100)

def calc_INCOME_STATEMENT_AB9(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AB6')), xl_ref(ctx.cell('INCOME STATEMENT', 'Z6'))), 1), 100)

def calc_INCOME_STATEMENT_B10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'B11')), xl_ref(ctx.cell('INCOME STATEMENT', 'B13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'B15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'B17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'B19')))

def calc_INCOME_STATEMENT_C10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'C11')), xl_ref(ctx.cell('INCOME STATEMENT', 'C13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'C15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'C17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'C19')))

def calc_INCOME_STATEMENT_E10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'E11')), xl_ref(ctx.cell('INCOME STATEMENT', 'E13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'E15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'E17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'E19')))

def calc_INCOME_STATEMENT_G10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'G11')), xl_ref(ctx.cell('INCOME STATEMENT', 'G13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'G15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'G17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'G19')))

def calc_INCOME_STATEMENT_I10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'I11')), xl_ref(ctx.cell('INCOME STATEMENT', 'I13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'I17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'I15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'I51')))

def calc_INCOME_STATEMENT_J10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'J11')), xl_ref(ctx.cell('INCOME STATEMENT', 'J13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'J17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'J15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'J19')))

def calc_INCOME_STATEMENT_L10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'L11')), xl_ref(ctx.cell('INCOME STATEMENT', 'L13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'L17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'L15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'L19')))

def calc_INCOME_STATEMENT_N10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'N11')), xl_ref(ctx.cell('INCOME STATEMENT', 'N13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'N17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'N15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'N19')))

def calc_INCOME_STATEMENT_P10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'P11')), xl_ref(ctx.cell('INCOME STATEMENT', 'P13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'P17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'P15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'P19')))

def calc_INCOME_STATEMENT_Q10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'Q11')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'U17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'Q15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'Q19')))

def calc_INCOME_STATEMENT_S10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'S11')), xl_ref(ctx.cell('INCOME STATEMENT', 'S13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'S15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'S17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'S19')))

def calc_INCOME_STATEMENT_U10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'U11')), xl_ref(ctx.cell('INCOME STATEMENT', 'U13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'U17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'U15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'U19')))

def calc_INCOME_STATEMENT_W10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'W11')), xl_ref(ctx.cell('INCOME STATEMENT', 'W13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'W17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'W15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'W19')))

def calc_INCOME_STATEMENT_X10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'X11')), xl_ref(ctx.cell('INCOME STATEMENT', 'X13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'X17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'X15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'X19')))

def calc_INCOME_STATEMENT_Z10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'Z11')), xl_ref(ctx.cell('INCOME STATEMENT', 'Z13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'Z17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'Z15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'Z19')))

def calc_INCOME_STATEMENT_AB10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AB11')), xl_ref(ctx.cell('INCOME STATEMENT', 'AB13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AB17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AB15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AB19')))

def calc_INCOME_STATEMENT_D11(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'B11')), xl_ref(ctx.cell('INCOME STATEMENT', 'C11')))

def calc_INCOME_STATEMENT_K11(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'I11')), xl_ref(ctx.cell('INCOME STATEMENT', 'J11')))

def calc_INCOME_STATEMENT_R11(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'P11')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q11')))

def calc_INCOME_STATEMENT_Y11(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'W11')), xl_ref(ctx.cell('INCOME STATEMENT', 'X11')))

def calc_INCOME_STATEMENT_AD11(ctx):
    return xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AD6')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AD12'))))

def calc_INCOME_STATEMENT_AE11(ctx):
    return xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AE6')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AE12'))))

def calc_INCOME_STATEMENT_AF11(ctx):
    return xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AF6')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AF12'))))

def calc_INCOME_STATEMENT_AG11(ctx):
    return xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AG6')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AG12'))))

def calc_INCOME_STATEMENT_B12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'B11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'B6'))))

def calc_INCOME_STATEMENT_C12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'C11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'C6'))))

def calc_INCOME_STATEMENT_E12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'E11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'E6'))))

def calc_INCOME_STATEMENT_G12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'G11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'G6'))))

def calc_INCOME_STATEMENT_H12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'H11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'H6'))))

def calc_INCOME_STATEMENT_I12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'I11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'I6'))))

def calc_INCOME_STATEMENT_J12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'J11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'J6'))))

def calc_INCOME_STATEMENT_L12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'L11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'L6'))))

def calc_INCOME_STATEMENT_N12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'N11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'N6'))))

def calc_INCOME_STATEMENT_P12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'P11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'P6'))))

def calc_INCOME_STATEMENT_Q12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Q11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'Q6'))))

def calc_INCOME_STATEMENT_S12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'S11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'S6'))))

def calc_INCOME_STATEMENT_U12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'U11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'U6'))))

def calc_INCOME_STATEMENT_V12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'V11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'V6'))))

def calc_INCOME_STATEMENT_W12(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'W11')), xl_ref(ctx.cell('INCOME STATEMENT', 'W6'))), 100)

def calc_INCOME_STATEMENT_X12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'X11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'X6'))))

def calc_INCOME_STATEMENT_Z12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Z11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'Z6'))))

def calc_INCOME_STATEMENT_AB12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AB11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AB6'))))

def calc_INCOME_STATEMENT_D13(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'B13')), xl_ref(ctx.cell('INCOME STATEMENT', 'C13')))

def calc_INCOME_STATEMENT_K13(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'I13')), xl_ref(ctx.cell('INCOME STATEMENT', 'J13')))

def calc_INCOME_STATEMENT_R13(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'P13')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q13')))

def calc_INCOME_STATEMENT_Y13(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'W13')), xl_ref(ctx.cell('INCOME STATEMENT', 'X13')))

def calc_INCOME_STATEMENT_AD13(ctx):
    return xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AD6')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AD14'))))

def calc_INCOME_STATEMENT_AE13(ctx):
    return xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AE6')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AE14'))))

def calc_INCOME_STATEMENT_AF13(ctx):
    return xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AF6')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AF14'))))

def calc_INCOME_STATEMENT_AG13(ctx):
    return xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AG6')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AG14'))))

def calc_INCOME_STATEMENT_B14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'B13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'B6'))))

def calc_INCOME_STATEMENT_C14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'C13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'C6'))))

def calc_INCOME_STATEMENT_E14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'E13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'E6'))))

def calc_INCOME_STATEMENT_G14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'G13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'G6'))))

def calc_INCOME_STATEMENT_H14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'H13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'H6'))))

def calc_INCOME_STATEMENT_I14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'I13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'I6'))))

def calc_INCOME_STATEMENT_J14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'J13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'J6'))))

def calc_INCOME_STATEMENT_L14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'L13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'L6'))))

def calc_INCOME_STATEMENT_N14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'N13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'N6'))))

def calc_INCOME_STATEMENT_P14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'P13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'P6'))))

def calc_INCOME_STATEMENT_Q14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Q13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'Q6'))))

def calc_INCOME_STATEMENT_S14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'S13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'S6'))))

def calc_INCOME_STATEMENT_U14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'U13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'U6'))))

def calc_INCOME_STATEMENT_V14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'V13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'V6'))))

def calc_INCOME_STATEMENT_W14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'W13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'W6'))))

def calc_INCOME_STATEMENT_X14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'X13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'X6'))))

def calc_INCOME_STATEMENT_Z14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Z13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'Z6'))))

def calc_INCOME_STATEMENT_AB14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AB13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AB6'))))

def calc_INCOME_STATEMENT_D15(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'B15')), xl_ref(ctx.cell('INCOME STATEMENT', 'C15')))

def calc_INCOME_STATEMENT_K15(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'I15')), xl_ref(ctx.cell('INCOME STATEMENT', 'J15')))

def calc_INCOME_STATEMENT_R15(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'P15')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q15')))

def calc_INCOME_STATEMENT_Y15(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'W15')), xl_ref(ctx.cell('INCOME STATEMENT', 'X15')))

def calc_INCOME_STATEMENT_AD15(ctx):
    return xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AD6')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AD16'))))

def calc_INCOME_STATEMENT_AE15(ctx):
    return xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AE6')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AE16'))))

def calc_INCOME_STATEMENT_AF15(ctx):
    return xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AF6')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AF16'))))

def calc_INCOME_STATEMENT_AG15(ctx):
    return xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AG6')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AG16'))))

def calc_INCOME_STATEMENT_B16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'B15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'B6'))))

def calc_INCOME_STATEMENT_C16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'C15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'C6'))))

def calc_INCOME_STATEMENT_E16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'E15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'E6'))))

def calc_INCOME_STATEMENT_G16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'G15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'G6'))))

def calc_INCOME_STATEMENT_I16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'I15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'I6'))))

def calc_INCOME_STATEMENT_J16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'J15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'J6'))))

def calc_INCOME_STATEMENT_L16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'L15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'L6'))))

def calc_INCOME_STATEMENT_N16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'N15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'N6'))))

def calc_INCOME_STATEMENT_P16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'P15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'P6'))))

def calc_INCOME_STATEMENT_Q16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Q15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'Q6'))))

def calc_INCOME_STATEMENT_S16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'S15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'S6'))))

def calc_INCOME_STATEMENT_U16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'U15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'U6'))))

def calc_INCOME_STATEMENT_W16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'W15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'W6'))))

def calc_INCOME_STATEMENT_X16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'X15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'X6'))))

def calc_INCOME_STATEMENT_Z16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Z15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'Z6'))))

def calc_INCOME_STATEMENT_AB16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AB15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AB6'))))

def calc_INCOME_STATEMENT_D17(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'B17')), xl_ref(ctx.cell('INCOME STATEMENT', 'C17')))

def calc_INCOME_STATEMENT_K17(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'I17')), xl_ref(ctx.cell('INCOME STATEMENT', 'J17')))

def calc_INCOME_STATEMENT_R17(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'P17')), xl_ref(ctx.cell('INCOME STATEMENT', 'U17')))

def calc_INCOME_STATEMENT_Y17(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'W17')), xl_ref(ctx.cell('INCOME STATEMENT', 'X17')))

def calc_INCOME_STATEMENT_AD17(ctx):
    return xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AD6')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AD18'))))

def calc_INCOME_STATEMENT_AE17(ctx):
    return xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AE6')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AE18'))))

def calc_INCOME_STATEMENT_AF17(ctx):
    return xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AF6')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AF18'))))

def calc_INCOME_STATEMENT_AG17(ctx):
    return xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AG6')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AG18'))))

def calc_INCOME_STATEMENT_B18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'B17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'B6'))))

def calc_INCOME_STATEMENT_C18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'C17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'C6'))))

def calc_INCOME_STATEMENT_E18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'E17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'E6'))))

def calc_INCOME_STATEMENT_G18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'G17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'G6'))))

def calc_INCOME_STATEMENT_H18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'H17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'H6'))))

def calc_INCOME_STATEMENT_I18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'I17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'I6'))))

def calc_INCOME_STATEMENT_J18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'J17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'J6'))))

def calc_INCOME_STATEMENT_L18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'L17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'L6'))))

def calc_INCOME_STATEMENT_N18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'N17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'N6'))))

def calc_INCOME_STATEMENT_P18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'P17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'P6'))))

def calc_INCOME_STATEMENT_Q18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'U17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'Q6'))))

def calc_INCOME_STATEMENT_S18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'S17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'S6'))))

def calc_INCOME_STATEMENT_U18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'U17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'U6'))))

def calc_INCOME_STATEMENT_V18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'V17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'V6'))))

def calc_INCOME_STATEMENT_W18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'W17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'W6'))))

def calc_INCOME_STATEMENT_X18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'X17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'X6'))))

def calc_INCOME_STATEMENT_Z18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Z17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'Z6'))))

def calc_INCOME_STATEMENT_AB18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AB17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AB6'))))

def calc_INCOME_STATEMENT_D19(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'B19')), xl_ref(ctx.cell('INCOME STATEMENT', 'C19')))

def calc_INCOME_STATEMENT_K19(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'I51')), xl_ref(ctx.cell('INCOME STATEMENT', 'J19')))

def calc_INCOME_STATEMENT_R19(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'P19')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q19')))

def calc_INCOME_STATEMENT_Y19(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'W19')), xl_ref(ctx.cell('INCOME STATEMENT', 'X19')))

def calc_INCOME_STATEMENT_AD19(ctx):
    return xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AD6')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AD20'))))

def calc_INCOME_STATEMENT_AE19(ctx):
    return xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AE6')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AE20'))))

def calc_INCOME_STATEMENT_AF19(ctx):
    return xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AF6')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AF20'))))

def calc_INCOME_STATEMENT_AG19(ctx):
    return xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AG6')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AG20'))))

def calc_INCOME_STATEMENT_B20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'B51')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'B6'))))

def calc_INCOME_STATEMENT_C20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'C51')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'C6'))))

def calc_INCOME_STATEMENT_E20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'E51')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'E6'))))

def calc_INCOME_STATEMENT_G20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'G19')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'G6'))))

def calc_INCOME_STATEMENT_I20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'I51')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'I6'))))

def calc_INCOME_STATEMENT_J20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'J19')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'J6'))))

def calc_INCOME_STATEMENT_L20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'L19')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'L6'))))

def calc_INCOME_STATEMENT_N20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'N19')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'N6'))))

def calc_INCOME_STATEMENT_P20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'P19')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'P6'))))

def calc_INCOME_STATEMENT_Q20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Q19')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'Q6'))))

def calc_INCOME_STATEMENT_S20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'S19')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'S6'))))

def calc_INCOME_STATEMENT_U20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'U19')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'U6'))))

def calc_INCOME_STATEMENT_W20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'W19')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'W6'))))

def calc_INCOME_STATEMENT_X20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'X19')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'X6'))))

def calc_INCOME_STATEMENT_Z20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Z19')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'Z6'))))

def calc_INCOME_STATEMENT_AB20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AB19')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AB6'))))

def calc_INCOME_STATEMENT_K21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'L6')), xl_ref(ctx.cell('INCOME STATEMENT', 'K10')))

def calc_INCOME_STATEMENT_D26(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'B26')), xl_ref(ctx.cell('INCOME STATEMENT', 'C26')))

def calc_INCOME_STATEMENT_K26(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'I26')), xl_ref(ctx.cell('INCOME STATEMENT', 'J26')))

def calc_INCOME_STATEMENT_R26(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'P26')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q26')))

def calc_INCOME_STATEMENT_Y26(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'W26')), xl_ref(ctx.cell('INCOME STATEMENT', 'X26')))

def calc_INCOME_STATEMENT_D27(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'B27')), xl_ref(ctx.cell('INCOME STATEMENT', 'C27')))

def calc_INCOME_STATEMENT_K27(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'I27')), xl_ref(ctx.cell('INCOME STATEMENT', 'J27')))

def calc_INCOME_STATEMENT_R27(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'P27')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q27')))

def calc_INCOME_STATEMENT_Y27(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'W27')), xl_ref(ctx.cell('INCOME STATEMENT', 'X27')))

def calc_INCOME_STATEMENT_D29(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'B29')), xl_ref(ctx.cell('INCOME STATEMENT', 'C29')))

def calc_INCOME_STATEMENT_K29(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'I29')), xl_ref(ctx.cell('INCOME STATEMENT', 'J29')))

def calc_INCOME_STATEMENT_R29(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'P29')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q29')))

def calc_INCOME_STATEMENT_V29(ctx):
    return xl_add(61177, 18485)

def calc_INCOME_STATEMENT_Y29(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'W29')), xl_ref(ctx.cell('INCOME STATEMENT', 'X29')))

def calc_INCOME_STATEMENT_M31(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'K31')), xl_ref(ctx.cell('INCOME STATEMENT', 'L31')))

def calc_INCOME_STATEMENT_R31(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'P31')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q31')))

def calc_INCOME_STATEMENT_Y31(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'W31')), xl_ref(ctx.cell('INCOME STATEMENT', 'X31')))

def calc_INCOME_STATEMENT_D33(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'B33')), xl_ref(ctx.cell('INCOME STATEMENT', 'C33')))

def calc_INCOME_STATEMENT_K33(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'I33')), xl_ref(ctx.cell('INCOME STATEMENT', 'J33')))

def calc_INCOME_STATEMENT_R33(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'P33')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q33')))

def calc_INCOME_STATEMENT_Y33(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'W33')), xl_ref(ctx.cell('INCOME STATEMENT', 'X33')))

def calc_INCOME_STATEMENT_D34(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'B34')), xl_ref(ctx.cell('INCOME STATEMENT', 'C34')))

def calc_INCOME_STATEMENT_K34(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'I34')), xl_ref(ctx.cell('INCOME STATEMENT', 'J34')))

def calc_INCOME_STATEMENT_R34(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'P34')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q34')))

def calc_INCOME_STATEMENT_Y34(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'W34')), xl_ref(ctx.cell('INCOME STATEMENT', 'X34')))

def calc_INCOME_STATEMENT_D35(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'B35')), xl_ref(ctx.cell('INCOME STATEMENT', 'C35')))

def calc_INCOME_STATEMENT_K35(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'I35')), xl_ref(ctx.cell('INCOME STATEMENT', 'J35')))

def calc_INCOME_STATEMENT_R35(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'P35')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q35')))

def calc_INCOME_STATEMENT_Y35(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'W35')), xl_ref(ctx.cell('INCOME STATEMENT', 'X35')))

def calc_INCOME_STATEMENT_D39(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'B39')), xl_ref(ctx.cell('INCOME STATEMENT', 'C39')))

def calc_INCOME_STATEMENT_H39(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'F39')), xl_ref(ctx.cell('INCOME STATEMENT', 'G39')))

def calc_INCOME_STATEMENT_D40(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'B40')), xl_ref(ctx.cell('INCOME STATEMENT', 'C40')))

def calc_INCOME_STATEMENT_AE43(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AE41')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AD41'))))

def calc_INCOME_STATEMENT_AF43(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AF41')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AE41'))))

def calc_INCOME_STATEMENT_AG43(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AG41')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AF41'))))

def calc_INCOME_STATEMENT_AD45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AD41')), xl_ref(ctx.cell('INCOME STATEMENT', 'AD6'))), 100)

def calc_INCOME_STATEMENT_AE45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AE41')), xl_ref(ctx.cell('INCOME STATEMENT', 'AE6'))), 100)

def calc_INCOME_STATEMENT_AF45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AF41')), xl_ref(ctx.cell('INCOME STATEMENT', 'AF6'))), 100)

def calc_INCOME_STATEMENT_AG45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AG41')), xl_ref(ctx.cell('INCOME STATEMENT', 'AG6'))), 100)

def calc_INCOME_STATEMENT_C47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'B47'))

def calc_INCOME_STATEMENT_T47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'S47'))

def calc_INCOME_STATEMENT_B48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'B47')), 10)

def calc_INCOME_STATEMENT_S48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'S47')), 10)

def calc_INCOME_STATEMENT_W66(ctx):
    return xl_sub(81, 58)

def calc_CASH_FOW_STATEMENT_I8(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AH32'))

def calc_CASH_FOW_STATEMENT_C9(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'O33')), xl_ref(ctx.cell('INCOME STATEMENT', 'O34')))

def calc_CASH_FOW_STATEMENT_D9(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'V33')), xl_ref(ctx.cell('INCOME STATEMENT', 'V34')))

def calc_CASH_FOW_STATEMENT_I9(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AH33')), xl_ref(ctx.cell('INCOME STATEMENT', 'AH34')))

def calc_CASH_FOW_STATEMENT_E11(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AC27'))

def calc_CASH_FOW_STATEMENT_F11(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AD27'))

def calc_CASH_FOW_STATEMENT_G11(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AE27'))

def calc_CASH_FOW_STATEMENT_H11(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AF27'))

def calc_CASH_FOW_STATEMENT_I11(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AG27'))

def calc_CASH_FOW_STATEMENT_B13(ctx):
    return xl_add(xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'C28')), xl_ref(ctx.cell('BALANCESHEET', 'B28'))), xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'C31')), xl_ref(ctx.cell('BALANCESHEET', 'B31'))))

def calc_CASH_FOW_STATEMENT_C13(ctx):
    return xl_add(xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'D28')), xl_ref(ctx.cell('BALANCESHEET', 'C28'))), xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'D31')), xl_ref(ctx.cell('BALANCESHEET', 'C31'))))

def calc_CASH_FOW_STATEMENT_D13(ctx):
    return xl_add(xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'E28')), xl_ref(ctx.cell('BALANCESHEET', 'D28'))), xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'E31')), xl_ref(ctx.cell('BALANCESHEET', 'D31'))))

def calc_CASH_FOW_STATEMENT_E13(ctx):
    return xl_add(xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'F28')), xl_ref(ctx.cell('BALANCESHEET', 'E28'))), xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'F31')), xl_ref(ctx.cell('BALANCESHEET', 'E31'))))

def calc_CASH_FOW_STATEMENT_F13(ctx):
    return xl_add(xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'G28')), xl_ref(ctx.cell('BALANCESHEET', 'F28'))), xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'G31')), xl_ref(ctx.cell('BALANCESHEET', 'F31'))))

def calc_CASH_FOW_STATEMENT_G13(ctx):
    return xl_add(xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'H28')), xl_ref(ctx.cell('BALANCESHEET', 'G28'))), xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'H31')), xl_ref(ctx.cell('BALANCESHEET', 'G31'))))

def calc_CASH_FOW_STATEMENT_H13(ctx):
    return xl_add(xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'I28')), xl_ref(ctx.cell('BALANCESHEET', 'H28'))), xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'I31')), xl_ref(ctx.cell('BALANCESHEET', 'H31'))))

def calc_CASH_FOW_STATEMENT_I13(ctx):
    return xl_add(xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'J28')), xl_ref(ctx.cell('BALANCESHEET', 'I28'))), xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'J31')), xl_ref(ctx.cell('BALANCESHEET', 'I31'))))

def calc_CASH_FOW_STATEMENT_B17(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'C9')), xl_ref(ctx.cell('BALANCESHEET', 'B9')))

def calc_CASH_FOW_STATEMENT_C17(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'D9')), xl_ref(ctx.cell('BALANCESHEET', 'C9')))

def calc_CASH_FOW_STATEMENT_D17(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'E9')), xl_ref(ctx.cell('BALANCESHEET', 'D9')))

def calc_CASH_FOW_STATEMENT_E17(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'F9')), xl_ref(ctx.cell('BALANCESHEET', 'E9')))

def calc_CASH_FOW_STATEMENT_F17(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'G9')), xl_ref(ctx.cell('BALANCESHEET', 'F9')))

def calc_CASH_FOW_STATEMENT_G17(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'H9')), xl_ref(ctx.cell('BALANCESHEET', 'G9')))

def calc_CASH_FOW_STATEMENT_H17(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'I9')), xl_ref(ctx.cell('BALANCESHEET', 'H9')))

def calc_CASH_FOW_STATEMENT_I17(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'J9')), xl_ref(ctx.cell('BALANCESHEET', 'I9')))

def calc_CASH_FOW_STATEMENT_B19(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'C33')), xl_ref(ctx.cell('BALANCESHEET', 'B33')))

def calc_CASH_FOW_STATEMENT_C19(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'D33')), xl_ref(ctx.cell('BALANCESHEET', 'C33')))

def calc_CASH_FOW_STATEMENT_D19(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'E33')), xl_ref(ctx.cell('BALANCESHEET', 'D33')))

def calc_CASH_FOW_STATEMENT_E19(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'F33')), xl_ref(ctx.cell('BALANCESHEET', 'E33')))

def calc_CASH_FOW_STATEMENT_F19(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'G33')), xl_ref(ctx.cell('BALANCESHEET', 'F33')))

def calc_CASH_FOW_STATEMENT_G19(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'H33')), xl_ref(ctx.cell('BALANCESHEET', 'G33')))

def calc_CASH_FOW_STATEMENT_H19(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'I33')), xl_ref(ctx.cell('BALANCESHEET', 'H33')))

def calc_CASH_FOW_STATEMENT_I19(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'J33')), xl_ref(ctx.cell('BALANCESHEET', 'I33')))

def calc_CASH_FOW_STATEMENT_F20(ctx):
    return xl_sub(xl_sub(xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'G37')), xl_ref(ctx.cell('BALANCESHEET', 'G38'))), xl_ref(ctx.cell('BALANCESHEET', 'G40'))), xl_ref(ctx.cell('BALANCESHEET', 'G41'))), xl_ref(ctx.cell('BALANCESHEET', 'G43'))), xl_sub(xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'F37')), xl_ref(ctx.cell('BALANCESHEET', 'F38'))), xl_ref(ctx.cell('BALANCESHEET', 'F40'))), xl_ref(ctx.cell('BALANCESHEET', 'F41'))), xl_ref(ctx.cell('BALANCESHEET', 'F43'))))

def calc_CASH_FOW_STATEMENT_G20(ctx):
    return xl_sub(xl_sub(xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'H37')), xl_ref(ctx.cell('BALANCESHEET', 'H38'))), xl_ref(ctx.cell('BALANCESHEET', 'H40'))), xl_ref(ctx.cell('BALANCESHEET', 'H41'))), xl_ref(ctx.cell('BALANCESHEET', 'H43'))), xl_sub(xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'G37')), xl_ref(ctx.cell('BALANCESHEET', 'G38'))), xl_ref(ctx.cell('BALANCESHEET', 'G40'))), xl_ref(ctx.cell('BALANCESHEET', 'G41'))), xl_ref(ctx.cell('BALANCESHEET', 'G43'))))

def calc_CASH_FOW_STATEMENT_H20(ctx):
    return xl_sub(xl_sub(xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'I37')), xl_ref(ctx.cell('BALANCESHEET', 'I38'))), xl_ref(ctx.cell('BALANCESHEET', 'I40'))), xl_ref(ctx.cell('BALANCESHEET', 'I41'))), xl_ref(ctx.cell('BALANCESHEET', 'I43'))), xl_sub(xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'H37')), xl_ref(ctx.cell('BALANCESHEET', 'H38'))), xl_ref(ctx.cell('BALANCESHEET', 'H40'))), xl_ref(ctx.cell('BALANCESHEET', 'H41'))), xl_ref(ctx.cell('BALANCESHEET', 'H43'))))

def calc_CASH_FOW_STATEMENT_I20(ctx):
    return xl_sub(xl_sub(xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'J37')), xl_ref(ctx.cell('BALANCESHEET', 'J38'))), xl_ref(ctx.cell('BALANCESHEET', 'J40'))), xl_ref(ctx.cell('BALANCESHEET', 'J41'))), xl_ref(ctx.cell('BALANCESHEET', 'J43'))), xl_sub(xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'I37')), xl_ref(ctx.cell('BALANCESHEET', 'I38'))), xl_ref(ctx.cell('BALANCESHEET', 'I40'))), xl_ref(ctx.cell('BALANCESHEET', 'I41'))), xl_ref(ctx.cell('BALANCESHEET', 'I43'))))

def calc_CASH_FOW_STATEMENT_E22(ctx):
    return xl_ref(ctx.cell('CASH FOW STATEMENT', 'D24'))

def calc_BALANCESHEET_B10(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'B11:B17'))

def calc_BALANCESHEET_C10(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'C11:C17'))

def calc_BALANCESHEET_D10(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'D11:D17'))

def calc_BALANCESHEET_E10(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'E11:E17'))

def calc_BALANCESHEET_F10(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'F11:F17'))

def calc_BALANCESHEET_G10(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'G11:G17'))

def calc_BALANCESHEET_H10(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'H11:H17'))

def calc_BALANCESHEET_I10(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'I11:I17'))

def calc_BALANCESHEET_J10(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'J11:J17'))

def calc_BALANCESHEET_B19(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'B20:B21'))

def calc_BALANCESHEET_C19(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'C20:C21'))

def calc_BALANCESHEET_D19(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'D20:D21'))

def calc_BALANCESHEET_E19(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'E20:E21'))

def calc_BALANCESHEET_F19(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'F20:F21'))

def calc_BALANCESHEET_G19(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'G20:G21'))

def calc_BALANCESHEET_H19(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'H20:H21'))

def calc_BALANCESHEET_I19(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'I20:I21'))

def calc_BALANCESHEET_J19(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'J20:J21'))

def calc_BALANCESHEET_B30(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'B28')), xl_ref(ctx.cell('BALANCESHEET', 'B29')))

def calc_BALANCESHEET_C30(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'C28')), xl_ref(ctx.cell('BALANCESHEET', 'C29')))

def calc_BALANCESHEET_D30(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'D28')), xl_ref(ctx.cell('BALANCESHEET', 'D29')))

def calc_BALANCESHEET_E30(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'E28')), xl_ref(ctx.cell('BALANCESHEET', 'E29')))

def calc_BALANCESHEET_F30(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'F28')), xl_ref(ctx.cell('BALANCESHEET', 'F29')))

def calc_BALANCESHEET_G30(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'G28')), xl_ref(ctx.cell('BALANCESHEET', 'G29')))

def calc_BALANCESHEET_H30(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'H28')), xl_ref(ctx.cell('BALANCESHEET', 'H29')))

def calc_BALANCESHEET_I30(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'I28')), xl_ref(ctx.cell('BALANCESHEET', 'I29')))

def calc_BALANCESHEET_J30(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'J28')), xl_ref(ctx.cell('BALANCESHEET', 'J29')))

def calc_BALANCESHEET_B36(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'B37:B41'))

def calc_BALANCESHEET_C36(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'C37:C41'))

def calc_BALANCESHEET_D36(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'D37:D41'))

def calc_BALANCESHEET_E36(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'E37:E41'))

def calc_BALANCESHEET_F36(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'F37:F41'))

def calc_BALANCESHEET_G36(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'G37:G41'))

def calc_BALANCESHEET_H36(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'H37:H41'))

def calc_BALANCESHEET_I36(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'I37:I41'))

def calc_BALANCESHEET_J36(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'J37:J41'))

def calc_BALANCESHEET_B44(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'B45:B50'))

def calc_BALANCESHEET_C44(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'C45:C50'))

def calc_BALANCESHEET_D44(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'D45:D50'))

def calc_BALANCESHEET_E44(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'E45:E50'))

def calc_BALANCESHEET_F44(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'F45:F50'))

def calc_BALANCESHEET_G44(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'G45:G50'))

def calc_BALANCESHEET_H44(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'H45:H50'))

def calc_BALANCESHEET_I44(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'I45:I50'))

def calc_BALANCESHEET_J44(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'J45:J50'))

def calc_BALANCESHEET_B52(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'B53:B56'))

def calc_BALANCESHEET_C52(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'C53:C56'))

def calc_BALANCESHEET_D52(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'D53:D56'))

def calc_BALANCESHEET_E52(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'E53:E56'))

def calc_Debt_Schedule_B11(ctx):
    return xl_add(xl_ref(ctx.cell('Debt Schedule', 'B7')), xl_ref(ctx.cell('Debt Schedule', 'B9')))

def calc_Debt_Schedule_C11(ctx):
    return xl_add(xl_ref(ctx.cell('Debt Schedule', 'C7')), xl_ref(ctx.cell('Debt Schedule', 'C9')))

def calc_Debt_Schedule_D11(ctx):
    return xl_add(xl_ref(ctx.cell('Debt Schedule', 'D7')), xl_ref(ctx.cell('Debt Schedule', 'D9')))

def calc_Debt_Schedule_E11(ctx):
    return xl_add(xl_ref(ctx.cell('Debt Schedule', 'E7')), xl_ref(ctx.cell('Debt Schedule', 'E9')))

def calc_Debt_Schedule_F11(ctx):
    return xl_add(xl_ref(ctx.cell('Debt Schedule', 'F7')), xl_ref(ctx.cell('Debt Schedule', 'F9')))

def calc_Debt_Schedule_G11(ctx):
    return xl_add(xl_ref(ctx.cell('Debt Schedule', 'G7')), xl_ref(ctx.cell('Debt Schedule', 'G9')))

def calc_Debt_Schedule_H11(ctx):
    return xl_add(xl_ref(ctx.cell('Debt Schedule', 'H7')), xl_ref(ctx.cell('Debt Schedule', 'H9')))

def calc_Debt_Schedule_I11(ctx):
    return xl_add(xl_ref(ctx.cell('Debt Schedule', 'I7')), xl_ref(ctx.cell('Debt Schedule', 'I9')))

def calc_Debt_Schedule_B20(ctx):
    return xl_add(xl_ref(ctx.cell('Debt Schedule', 'B16')), xl_ref(ctx.cell('Debt Schedule', 'B18')))

def calc_Debt_Schedule_C20(ctx):
    return xl_add(xl_ref(ctx.cell('Debt Schedule', 'C16')), xl_ref(ctx.cell('Debt Schedule', 'C18')))

def calc_Debt_Schedule_D20(ctx):
    return xl_add(xl_ref(ctx.cell('Debt Schedule', 'D16')), xl_ref(ctx.cell('Debt Schedule', 'D18')))

def calc_Debt_Schedule_E20(ctx):
    return xl_add(xl_ref(ctx.cell('Debt Schedule', 'E16')), xl_ref(ctx.cell('Debt Schedule', 'E18')))

def calc_Debt_Schedule_F20(ctx):
    return xl_add(xl_ref(ctx.cell('Debt Schedule', 'F16')), xl_ref(ctx.cell('Debt Schedule', 'F18')))

def calc_Debt_Schedule_G20(ctx):
    return xl_add(xl_ref(ctx.cell('Debt Schedule', 'G16')), xl_ref(ctx.cell('Debt Schedule', 'G18')))

def calc_Debt_Schedule_H20(ctx):
    return xl_add(xl_ref(ctx.cell('Debt Schedule', 'H16')), xl_ref(ctx.cell('Debt Schedule', 'H18')))

def calc_Debt_Schedule_I20(ctx):
    return xl_add(xl_ref(ctx.cell('Debt Schedule', 'I16')), xl_ref(ctx.cell('Debt Schedule', 'I18')))

def calc_Valuation_G9(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'G8')), xl_ref(ctx.cell('Valuation', 'G7')))

def calc_Valuation_B10(ctx):
    return xl_add(xl_ref(ctx.cell('Valuation', 'B11')), xl_mul(xl_ref(ctx.cell('Valuation', 'B12')), xl_ref(ctx.cell('Valuation', 'B13'))))

def calc_Valuation_G11(ctx):
    return xl_ref(ctx.cell('BALANCESHEET', 'E39'))

def calc_Valuation_B19(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'V6'))

def calc_Valuation_D19(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AD6'))

def calc_Valuation_E19(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AE6'))

def calc_Valuation_F19(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AF6'))

def calc_Valuation_G19(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AG6'))

def calc_Valuation_B22(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'V33'))

def calc_Valuation_C24(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AC27'))

def calc_Valuation_D24(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AD27'))

def calc_Valuation_E24(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AE27'))

def calc_Valuation_F24(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AF27'))

def calc_Valuation_G24(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AG27'))

def calc_Valuation_K37(ctx):
    return xl_ref(ctx.cell('Valuation', 'B6'))

def calc_Valuation_K53(ctx):
    return xl_ref(ctx.cell('Valuation', 'G8'))

def calc_Valuation_B54(ctx):
    return xl_ref(ctx.cell('Valuation', 'B6'))

def calc_Valuation_H64(ctx):
    return xl_ref(ctx.cell('Valuation', 'D64'))

def calc_Valuation_I64(ctx):
    return xl_ref(ctx.cell('Valuation', 'E64'))

def calc_Valuation_J64(ctx):
    return xl_ref(ctx.cell('Valuation', 'F64'))

def calc_Valuation_H74(ctx):
    return xl_ref(ctx.cell('Valuation', 'D74'))

def calc_Valuation_I74(ctx):
    return xl_ref(ctx.cell('Valuation', 'E74'))

def calc_Valuation_J74(ctx):
    return xl_ref(ctx.cell('Valuation', 'F74'))

def calc_Ratio_Analysis_G14(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AD41')), 100), xl_ref(ctx.cell('INCOME STATEMENT', 'AD6')))

def calc_Ratio_Analysis_H14(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AE41')), 100), xl_ref(ctx.cell('INCOME STATEMENT', 'AE6')))

def calc_Ratio_Analysis_I14(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AF41')), 100), xl_ref(ctx.cell('INCOME STATEMENT', 'AF6')))

def calc_Ratio_Analysis_J14(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AG41')), 100), xl_ref(ctx.cell('INCOME STATEMENT', 'AG6')))

def calc_Ratio_Analysis_C20(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('BALANCESHEET', 'C28')), xl_ref(ctx.cell('BALANCESHEET', 'B28'))), 1)

def calc_Ratio_Analysis_D20(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('BALANCESHEET', 'D28')), xl_ref(ctx.cell('BALANCESHEET', 'C28'))), 1)

def calc_Ratio_Analysis_E20(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('BALANCESHEET', 'E28')), xl_ref(ctx.cell('BALANCESHEET', 'D28'))), 1)

def calc_Ratio_Analysis_F20(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('BALANCESHEET', 'F28')), xl_ref(ctx.cell('BALANCESHEET', 'E28'))), 1)

def calc_Ratio_Analysis_G20(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('BALANCESHEET', 'G28')), xl_ref(ctx.cell('BALANCESHEET', 'F28'))), 1)

def calc_Ratio_Analysis_H20(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('BALANCESHEET', 'H28')), xl_ref(ctx.cell('BALANCESHEET', 'G28'))), 1)

def calc_Ratio_Analysis_I20(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('BALANCESHEET', 'I28')), xl_ref(ctx.cell('BALANCESHEET', 'H28'))), 1)

def calc_Ratio_Analysis_J20(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('BALANCESHEET', 'J28')), xl_ref(ctx.cell('BALANCESHEET', 'I28'))), 1)

def calc_Ratio_Analysis_G21(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AD6')), xl_ref(ctx.cell('INCOME STATEMENT', 'W6'))), 1)

def calc_Ratio_Analysis_H21(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AE6')), xl_ref(ctx.cell('INCOME STATEMENT', 'X6'))), 1)

def calc_Ratio_Analysis_J21(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AG6')), xl_ref(ctx.cell('INCOME STATEMENT', 'Z6'))), 1)

def calc_Ratio_Analysis_E23(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'V6')), xl_mul(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'E37')), xl_ref(ctx.cell('BALANCESHEET', 'D37'))), 0.5))

def calc_Ratio_Analysis_G23(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AD6')), xl_mul(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'G37')), xl_ref(ctx.cell('BALANCESHEET', 'F37'))), 0.5))

def calc_Ratio_Analysis_H23(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AE6')), xl_mul(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'H37')), xl_ref(ctx.cell('BALANCESHEET', 'G37'))), 0.5))

def calc_Ratio_Analysis_I23(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AF6')), xl_mul(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'I37')), xl_ref(ctx.cell('BALANCESHEET', 'H37'))), 0.5))

def calc_Ratio_Analysis_J23(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AG6')), xl_mul(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'J37')), xl_ref(ctx.cell('BALANCESHEET', 'I37'))), 0.5))

def calc_Ratio_Analysis_G24(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'G58')), xl_ref(ctx.cell('INCOME STATEMENT', 'AD6')))

def calc_Ratio_Analysis_H24(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'H58')), xl_ref(ctx.cell('INCOME STATEMENT', 'AE6')))

def calc_Ratio_Analysis_I24(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'I58')), xl_ref(ctx.cell('INCOME STATEMENT', 'AF6')))

def calc_Ratio_Analysis_J24(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'J58')), xl_ref(ctx.cell('INCOME STATEMENT', 'AG6')))

def calc_PRESENTATION_F15(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'F8')), xl_ref(ctx.cell('PRESENTATION', 'F9'))), xl_ref(ctx.cell('PRESENTATION', 'F10'))), xl_ref(ctx.cell('PRESENTATION', 'F11'))), xl_ref(ctx.cell('PRESENTATION', 'F12'))), xl_ref(ctx.cell('PRESENTATION', 'F13'))), xl_ref(ctx.cell('PRESENTATION', 'F14')))

def calc_PRESENTATION_K15(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'K8')), xl_ref(ctx.cell('PRESENTATION', 'K9'))), xl_ref(ctx.cell('PRESENTATION', 'K10'))), xl_ref(ctx.cell('PRESENTATION', 'K11'))), xl_ref(ctx.cell('PRESENTATION', 'K12'))), xl_ref(ctx.cell('PRESENTATION', 'K13'))), xl_ref(ctx.cell('PRESENTATION', 'K14')))

def calc_PRESENTATION_P15(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'P8')), xl_ref(ctx.cell('PRESENTATION', 'P9'))), xl_ref(ctx.cell('PRESENTATION', 'P10'))), xl_ref(ctx.cell('PRESENTATION', 'P11'))), xl_ref(ctx.cell('PRESENTATION', 'P12'))), xl_ref(ctx.cell('PRESENTATION', 'P13'))), xl_ref(ctx.cell('PRESENTATION', 'P14')))

def calc_PRESENTATION_T15(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'T8:T14'))

def calc_PRESENTATION_B17(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'B15')), xl_ref(ctx.cell('PRESENTATION', 'B16')))

def calc_PRESENTATION_C17(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'C15')), xl_ref(ctx.cell('PRESENTATION', 'C16')))

def calc_PRESENTATION_D17(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'D15')), xl_ref(ctx.cell('PRESENTATION', 'D16')))

def calc_PRESENTATION_E17(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'E15')), xl_ref(ctx.cell('PRESENTATION', 'E16')))

def calc_PRESENTATION_G17(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'G15')), xl_ref(ctx.cell('PRESENTATION', 'G16')))

def calc_PRESENTATION_H17(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'H15')), xl_ref(ctx.cell('PRESENTATION', 'H16')))

def calc_PRESENTATION_I17(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'I15')), xl_ref(ctx.cell('PRESENTATION', 'I16')))

def calc_PRESENTATION_J17(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'J15')), xl_ref(ctx.cell('PRESENTATION', 'J16')))

def calc_PRESENTATION_L17(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'L15')), xl_ref(ctx.cell('PRESENTATION', 'L16')))

def calc_PRESENTATION_M17(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'M15')), xl_ref(ctx.cell('PRESENTATION', 'M16')))

def calc_PRESENTATION_N17(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'N15')), xl_ref(ctx.cell('PRESENTATION', 'N16')))

def calc_PRESENTATION_O17(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'O15')), xl_ref(ctx.cell('PRESENTATION', 'O16')))

def calc_PRESENTATION_Q17(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'Q15:Q16'))

def calc_PRESENTATION_R17(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'R15:R16'))

def calc_PRESENTATION_S17(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'S15:S16'))

def calc_PRESENTATION_V17(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'V15')), xl_ref(ctx.cell('PRESENTATION', 'V16')))

def calc_PRESENTATION_W17(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'W15')), xl_ref(ctx.cell('PRESENTATION', 'W16')))

def calc_PRESENTATION_X17(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'X15')), xl_ref(ctx.cell('PRESENTATION', 'X16')))

def calc_PRESENTATION_Y17(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'Y15')), xl_ref(ctx.cell('PRESENTATION', 'Y16')))

def calc_PRESENTATION_F25(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'D25')), xl_ref(ctx.cell('PRESENTATION', 'E25')))

def calc_PRESENTATION_M25(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'K25')), xl_ref(ctx.cell('PRESENTATION', 'L25')))

def calc_PRESENTATION_T25(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'R25')), xl_ref(ctx.cell('PRESENTATION', 'S25')))

def calc_PRESENTATION_AA25(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'Y25')), xl_ref(ctx.cell('PRESENTATION', 'Z25')))

def calc_PRESENTATION_B32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'B25')), xl_ref(ctx.cell('PRESENTATION', 'B26')))

def calc_PRESENTATION_C32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'C25')), xl_ref(ctx.cell('PRESENTATION', 'C26')))

def calc_PRESENTATION_E32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'E25')), xl_ref(ctx.cell('PRESENTATION', 'E26')))

def calc_PRESENTATION_G32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'G25')), xl_ref(ctx.cell('PRESENTATION', 'G26')))

def calc_PRESENTATION_I32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'I25')), xl_ref(ctx.cell('PRESENTATION', 'I26')))

def calc_PRESENTATION_J32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'J25')), xl_ref(ctx.cell('PRESENTATION', 'J26')))

def calc_PRESENTATION_L32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'L25')), xl_ref(ctx.cell('PRESENTATION', 'L26')))

def calc_PRESENTATION_N32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'N25')), xl_ref(ctx.cell('PRESENTATION', 'N26')))

def calc_PRESENTATION_P32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'P25')), xl_ref(ctx.cell('PRESENTATION', 'P26')))

def calc_PRESENTATION_Q32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'Q25')), xl_ref(ctx.cell('PRESENTATION', 'Q26')))

def calc_PRESENTATION_S32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'S25')), xl_ref(ctx.cell('PRESENTATION', 'S26')))

def calc_PRESENTATION_U32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'U25')), xl_ref(ctx.cell('PRESENTATION', 'U26')))

def calc_PRESENTATION_W32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'W25')), xl_ref(ctx.cell('PRESENTATION', 'W26')))

def calc_PRESENTATION_X32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'X25')), xl_ref(ctx.cell('PRESENTATION', 'X26')))

def calc_PRESENTATION_Z32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'Z25')), xl_ref(ctx.cell('PRESENTATION', 'Z26')))

def calc_PRESENTATION_AB32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AB25')), xl_ref(ctx.cell('PRESENTATION', 'AB26')))

def calc_PRESENTATION_AD32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AD25')), xl_ref(ctx.cell('PRESENTATION', 'AD26')))

def calc_PRESENTATION_AE32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AE25')), xl_ref(ctx.cell('PRESENTATION', 'AE26')))

def calc_PRESENTATION_AF32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AF25')), xl_ref(ctx.cell('PRESENTATION', 'AF26')))

def calc_PRESENTATION_AG32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AG25')), xl_ref(ctx.cell('PRESENTATION', 'AG26')))

def calc_PRESENTATION_F27(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'D27')), xl_ref(ctx.cell('PRESENTATION', 'E27')))

def calc_PRESENTATION_M27(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'K27')), xl_ref(ctx.cell('PRESENTATION', 'L27')))

def calc_PRESENTATION_T27(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'R27')), xl_ref(ctx.cell('PRESENTATION', 'S27')))

def calc_PRESENTATION_AA27(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'Y27')), xl_ref(ctx.cell('PRESENTATION', 'Z27')))

def calc_PRESENTATION_F28(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'E28')), xl_ref(ctx.cell('PRESENTATION', 'D28')))

def calc_PRESENTATION_M28(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'L28')), xl_ref(ctx.cell('PRESENTATION', 'K28')))

def calc_PRESENTATION_T28(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'S28')), xl_ref(ctx.cell('PRESENTATION', 'R28')))

def calc_PRESENTATION_AA28(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'Z28')), xl_ref(ctx.cell('PRESENTATION', 'Y28')))

def calc_PRESENTATION_F29(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'E29')), xl_ref(ctx.cell('PRESENTATION', 'D29')))

def calc_PRESENTATION_M29(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'L29')), xl_ref(ctx.cell('PRESENTATION', 'K29')))

def calc_PRESENTATION_T29(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'S29')), xl_ref(ctx.cell('PRESENTATION', 'R29')))

def calc_PRESENTATION_AA29(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'Z29')), xl_ref(ctx.cell('PRESENTATION', 'Y29')))

def calc_PRESENTATION_F30(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'E30')), xl_ref(ctx.cell('PRESENTATION', 'D30')))

def calc_PRESENTATION_M30(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'L30')), xl_ref(ctx.cell('PRESENTATION', 'K30')))

def calc_PRESENTATION_T30(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'S30')), xl_ref(ctx.cell('PRESENTATION', 'R30')))

def calc_PRESENTATION_AA30(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'Z30')), xl_ref(ctx.cell('PRESENTATION', 'Y30')))

def calc_PRESENTATION_D26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'D27')), xl_ref(ctx.cell('PRESENTATION', 'D28'))), xl_ref(ctx.cell('PRESENTATION', 'D29'))), xl_ref(ctx.cell('PRESENTATION', 'D30'))), xl_ref(ctx.cell('PRESENTATION', 'D31')))

def calc_PRESENTATION_F31(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'E31')), xl_ref(ctx.cell('PRESENTATION', 'D31')))

def calc_PRESENTATION_M31(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'L31')), xl_ref(ctx.cell('PRESENTATION', 'K31')))

def calc_PRESENTATION_R26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'R27')), xl_ref(ctx.cell('PRESENTATION', 'R28'))), xl_ref(ctx.cell('PRESENTATION', 'R30'))), xl_ref(ctx.cell('PRESENTATION', 'R29'))), xl_ref(ctx.cell('PRESENTATION', 'R31')))

def calc_PRESENTATION_T31(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'S31')), xl_ref(ctx.cell('PRESENTATION', 'R31')))

def calc_PRESENTATION_Y26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'Y27')), xl_ref(ctx.cell('PRESENTATION', 'Y28'))), xl_ref(ctx.cell('PRESENTATION', 'Y30'))), xl_ref(ctx.cell('PRESENTATION', 'Y29'))), xl_ref(ctx.cell('PRESENTATION', 'Y31')))

def calc_PRESENTATION_AA31(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'Z31')), xl_ref(ctx.cell('PRESENTATION', 'Y31')))

def calc_PRESENTATION_K33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'K32')), xl_ref(ctx.cell('PRESENTATION', 'L25'))), 100)

def calc_PRESENTATION_F34(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'D34')), xl_ref(ctx.cell('PRESENTATION', 'E34')))

def calc_PRESENTATION_M34(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'K34')), xl_ref(ctx.cell('PRESENTATION', 'L34')))

def calc_PRESENTATION_T34(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'R34')), xl_ref(ctx.cell('PRESENTATION', 'S34')))

def calc_PRESENTATION_AA34(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'Y34')), xl_ref(ctx.cell('PRESENTATION', 'Z34')))

def calc_PRESENTATION_F35(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'D35')), xl_ref(ctx.cell('PRESENTATION', 'E35')))

def calc_PRESENTATION_M35(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'K35')), xl_ref(ctx.cell('PRESENTATION', 'L35')))

def calc_PRESENTATION_K36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'K32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'K34')), xl_ref(ctx.cell('PRESENTATION', 'K35'))))

def calc_PRESENTATION_T35(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'R35')), xl_ref(ctx.cell('PRESENTATION', 'S35')))

def calc_PRESENTATION_AA35(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'Y35')), xl_ref(ctx.cell('PRESENTATION', 'Z35')))

def calc_PRESENTATION_F37(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'D37')), xl_ref(ctx.cell('PRESENTATION', 'E37')))

def calc_PRESENTATION_M37(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'K37')), xl_ref(ctx.cell('PRESENTATION', 'L37')))

def calc_PRESENTATION_T37(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'R37')), xl_ref(ctx.cell('PRESENTATION', 'S37')))

def calc_PRESENTATION_AA37(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'Y37')), xl_ref(ctx.cell('PRESENTATION', 'Z37')))

def calc_PRESENTATION_T39(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'R39')), xl_ref(ctx.cell('PRESENTATION', 'S39')))

def calc_PRESENTATION_AA39(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'Y39')), xl_ref(ctx.cell('PRESENTATION', 'Z39')))

def calc_PRESENTATION_F41(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'D41')), xl_ref(ctx.cell('PRESENTATION', 'E41')))

def calc_PRESENTATION_M41(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'K41')), xl_ref(ctx.cell('PRESENTATION', 'L41')))

def calc_PRESENTATION_T41(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'R41')), xl_ref(ctx.cell('PRESENTATION', 'S41')))

def calc_PRESENTATION_AA41(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'Y41')), xl_ref(ctx.cell('PRESENTATION', 'Z41')))

def calc_PRESENTATION_F42(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'D42')), xl_ref(ctx.cell('PRESENTATION', 'E42')))

def calc_PRESENTATION_M42(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'K42')), xl_ref(ctx.cell('PRESENTATION', 'L42')))

def calc_PRESENTATION_T42(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'R42')), xl_ref(ctx.cell('PRESENTATION', 'S42')))

def calc_PRESENTATION_AA42(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'Y42')), xl_ref(ctx.cell('PRESENTATION', 'Z42')))

def calc_PRESENTATION_F43(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'D43')), xl_ref(ctx.cell('PRESENTATION', 'E43')))

def calc_PRESENTATION_M43(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'K43')), xl_ref(ctx.cell('PRESENTATION', 'L43')))

def calc_PRESENTATION_T43(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'R43')), xl_ref(ctx.cell('PRESENTATION', 'S43')))

def calc_PRESENTATION_AA43(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'Y43')), xl_ref(ctx.cell('PRESENTATION', 'Z43')))

def calc_Segment_Revenue_Model_K24(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'K8')), xl_ref(ctx.cell('Segment Revenue Model', 'F8'))), 1)

def calc_Segment_Revenue_Model_U8(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'U24')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'P8')))

def calc_Segment_Revenue_Model_P24(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'P8')), xl_ref(ctx.cell('Segment Revenue Model', 'K8'))), 1)

def calc_Segment_Revenue_Model_K25(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'K9')), xl_ref(ctx.cell('Segment Revenue Model', 'F9'))), 1)

def calc_Segment_Revenue_Model_U9(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'U25')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'P9')))

def calc_Segment_Revenue_Model_P25(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'P9')), xl_ref(ctx.cell('Segment Revenue Model', 'K9'))), 1)

def calc_Segment_Revenue_Model_K26(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'K10')), xl_ref(ctx.cell('Segment Revenue Model', 'F10'))), 1)

def calc_Segment_Revenue_Model_U10(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'U26')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'P10')))

def calc_Segment_Revenue_Model_P26(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'P10')), xl_ref(ctx.cell('Segment Revenue Model', 'K10'))), 1)

def calc_Segment_Revenue_Model_K27(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'K11')), xl_ref(ctx.cell('Segment Revenue Model', 'F11'))), 1)

def calc_Segment_Revenue_Model_U11(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'U27')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'P11')))

def calc_Segment_Revenue_Model_P27(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'P11')), xl_ref(ctx.cell('Segment Revenue Model', 'K11'))), 1)

def calc_Segment_Revenue_Model_K28(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'K12')), xl_ref(ctx.cell('Segment Revenue Model', 'F12'))), 1)

def calc_Segment_Revenue_Model_P28(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'P12')), xl_ref(ctx.cell('Segment Revenue Model', 'K12'))), 1)

def calc_Segment_Revenue_Model_T12(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'U12')), xl_add(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'Q12')), xl_ref(ctx.cell('Segment Revenue Model', 'R12'))), xl_ref(ctx.cell('Segment Revenue Model', 'S12'))))

def calc_Segment_Revenue_Model_V12(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'V28')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'U12')))

def calc_Segment_Revenue_Model_K29(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'K13')), xl_ref(ctx.cell('Segment Revenue Model', 'F13'))), 1)

def calc_Segment_Revenue_Model_U13(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'U29')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'P13')))

def calc_Segment_Revenue_Model_P29(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'P13')), xl_ref(ctx.cell('Segment Revenue Model', 'K13'))), 1)

def calc_Segment_Revenue_Model_F15(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'F8')), xl_ref(ctx.cell('Segment Revenue Model', 'F9'))), xl_ref(ctx.cell('Segment Revenue Model', 'F10'))), xl_ref(ctx.cell('Segment Revenue Model', 'F11'))), xl_ref(ctx.cell('Segment Revenue Model', 'F12'))), xl_ref(ctx.cell('Segment Revenue Model', 'F13'))), xl_ref(ctx.cell('Segment Revenue Model', 'F14')))

def calc_Segment_Revenue_Model_K15(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'K8')), xl_ref(ctx.cell('Segment Revenue Model', 'K9'))), xl_ref(ctx.cell('Segment Revenue Model', 'K10'))), xl_ref(ctx.cell('Segment Revenue Model', 'K11'))), xl_ref(ctx.cell('Segment Revenue Model', 'K12'))), xl_ref(ctx.cell('Segment Revenue Model', 'K13'))), xl_ref(ctx.cell('Segment Revenue Model', 'K14')))

def calc_Segment_Revenue_Model_K30(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'K14')), xl_ref(ctx.cell('Segment Revenue Model', 'F14'))), 1)

def calc_Segment_Revenue_Model_U14(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'U30')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'P14')))

def calc_Segment_Revenue_Model_P15(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'P8')), xl_ref(ctx.cell('Segment Revenue Model', 'P9'))), xl_ref(ctx.cell('Segment Revenue Model', 'P10'))), xl_ref(ctx.cell('Segment Revenue Model', 'P11'))), xl_ref(ctx.cell('Segment Revenue Model', 'P12'))), xl_ref(ctx.cell('Segment Revenue Model', 'P13'))), xl_ref(ctx.cell('Segment Revenue Model', 'P14')))

def calc_Segment_Revenue_Model_P30(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'P14')), xl_ref(ctx.cell('Segment Revenue Model', 'K14'))), 1)

def calc_Segment_Revenue_Model_B17(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'B15')), xl_ref(ctx.cell('Segment Revenue Model', 'B16')))

def calc_Segment_Revenue_Model_C17(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'C15')), xl_ref(ctx.cell('Segment Revenue Model', 'C16')))

def calc_Segment_Revenue_Model_D17(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'D15')), xl_ref(ctx.cell('Segment Revenue Model', 'D16')))

def calc_Segment_Revenue_Model_E17(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'E15')), xl_ref(ctx.cell('Segment Revenue Model', 'E16')))

def calc_Segment_Revenue_Model_G17(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'G15')), xl_ref(ctx.cell('Segment Revenue Model', 'G16')))

def calc_Segment_Revenue_Model_G31(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'G15')), xl_ref(ctx.cell('Segment Revenue Model', 'B15'))), 1)

def calc_Segment_Revenue_Model_H17(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'H15')), xl_ref(ctx.cell('Segment Revenue Model', 'H16')))

def calc_Segment_Revenue_Model_H31(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'H15')), xl_ref(ctx.cell('Segment Revenue Model', 'C15'))), 1)

def calc_Segment_Revenue_Model_I17(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'I15')), xl_ref(ctx.cell('Segment Revenue Model', 'I16')))

def calc_Segment_Revenue_Model_I31(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'I15')), xl_ref(ctx.cell('Segment Revenue Model', 'D15'))), 1)

def calc_Segment_Revenue_Model_J17(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'J15')), xl_ref(ctx.cell('Segment Revenue Model', 'J16')))

def calc_Segment_Revenue_Model_J31(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'J15')), xl_ref(ctx.cell('Segment Revenue Model', 'E15'))), 1)

def calc_Segment_Revenue_Model_L17(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'L15')), xl_ref(ctx.cell('Segment Revenue Model', 'L16')))

def calc_Segment_Revenue_Model_L31(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'L15')), xl_ref(ctx.cell('Segment Revenue Model', 'G15'))), 1)

def calc_Segment_Revenue_Model_M17(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'M15')), xl_ref(ctx.cell('Segment Revenue Model', 'M16')))

def calc_Segment_Revenue_Model_M31(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'M15')), xl_ref(ctx.cell('Segment Revenue Model', 'H15'))), 1)

def calc_Segment_Revenue_Model_N17(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'N15')), xl_ref(ctx.cell('Segment Revenue Model', 'N16')))

def calc_Segment_Revenue_Model_N31(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'N15')), xl_ref(ctx.cell('Segment Revenue Model', 'I15'))), 1)

def calc_Segment_Revenue_Model_O17(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'O15')), xl_ref(ctx.cell('Segment Revenue Model', 'O16')))

def calc_Segment_Revenue_Model_O31(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'O15')), xl_ref(ctx.cell('Segment Revenue Model', 'J15'))), 1)

def calc_Segment_Revenue_Model_Q17(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'Q15:Q16'))

def calc_Segment_Revenue_Model_Q31(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Q15')), xl_ref(ctx.cell('Segment Revenue Model', 'L15'))), 1)

def calc_Segment_Revenue_Model_R17(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'R15:R16'))

def calc_Segment_Revenue_Model_R31(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'R15')), xl_ref(ctx.cell('Segment Revenue Model', 'M15'))), 1)

def calc_Segment_Revenue_Model_S17(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'S15:S16'))

def calc_Segment_Revenue_Model_S31(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'S15')), xl_ref(ctx.cell('Segment Revenue Model', 'N15'))), 1)

def calc_Segment_Revenue_Model_K32(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'K16')), xl_ref(ctx.cell('Segment Revenue Model', 'F16'))), 1)

def calc_Segment_Revenue_Model_P32(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'P16')), xl_ref(ctx.cell('Segment Revenue Model', 'K16'))), 1)

def calc_Segment_Revenue_Model_V16(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'V32')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'U16')))

def calc_Segment_Revenue_Model_F58(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'F51')), xl_ref(ctx.cell('Segment Revenue Model', 'F52'))), xl_ref(ctx.cell('Segment Revenue Model', 'F53'))), xl_ref(ctx.cell('Segment Revenue Model', 'F54'))), xl_ref(ctx.cell('Segment Revenue Model', 'F55'))), xl_ref(ctx.cell('Segment Revenue Model', 'F56'))), xl_ref(ctx.cell('Segment Revenue Model', 'F57')))

def calc_Segment_Revenue_Model_K58(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'K51')), xl_ref(ctx.cell('Segment Revenue Model', 'K53'))), xl_ref(ctx.cell('Segment Revenue Model', 'K54'))), xl_ref(ctx.cell('Segment Revenue Model', 'K55'))), xl_ref(ctx.cell('Segment Revenue Model', 'K56'))), xl_ref(ctx.cell('Segment Revenue Model', 'K57')))

def calc_Segment_Revenue_Model_P58(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'P51')), xl_ref(ctx.cell('Segment Revenue Model', 'P52'))), xl_ref(ctx.cell('Segment Revenue Model', 'P53'))), xl_ref(ctx.cell('Segment Revenue Model', 'P54'))), xl_ref(ctx.cell('Segment Revenue Model', 'P55'))), xl_ref(ctx.cell('Segment Revenue Model', 'P56'))), xl_ref(ctx.cell('Segment Revenue Model', 'P57')))

def calc_Segment_Revenue_Model_U58(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'U51')), xl_ref(ctx.cell('Segment Revenue Model', 'U52'))), xl_ref(ctx.cell('Segment Revenue Model', 'U53'))), xl_ref(ctx.cell('Segment Revenue Model', 'U54'))), xl_ref(ctx.cell('Segment Revenue Model', 'U55'))), xl_ref(ctx.cell('Segment Revenue Model', 'U56'))), xl_ref(ctx.cell('Segment Revenue Model', 'U57')))

def calc_Segment_Revenue_Model_B61(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'B58')), xl_ref(ctx.cell('Segment Revenue Model', 'B59'))), xl_ref(ctx.cell('Segment Revenue Model', 'B60')))

def calc_Segment_Revenue_Model_C61(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'C58')), xl_ref(ctx.cell('Segment Revenue Model', 'C59'))), xl_ref(ctx.cell('Segment Revenue Model', 'C60')))

def calc_Segment_Revenue_Model_D61(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'D58')), xl_ref(ctx.cell('Segment Revenue Model', 'D59'))), xl_ref(ctx.cell('Segment Revenue Model', 'D60')))

def calc_Segment_Revenue_Model_E61(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'E58')), xl_ref(ctx.cell('Segment Revenue Model', 'E59'))), xl_ref(ctx.cell('Segment Revenue Model', 'E60')))

def calc_Segment_Revenue_Model_G61(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'G58')), xl_ref(ctx.cell('Segment Revenue Model', 'G59'))), xl_ref(ctx.cell('Segment Revenue Model', 'G60')))

def calc_Segment_Revenue_Model_H61(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'H58')), xl_ref(ctx.cell('Segment Revenue Model', 'H59'))), xl_ref(ctx.cell('Segment Revenue Model', 'H60')))

def calc_Segment_Revenue_Model_I61(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'I58:I60'))

def calc_Segment_Revenue_Model_J61(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'J58')), xl_ref(ctx.cell('Segment Revenue Model', 'J59'))), xl_ref(ctx.cell('Segment Revenue Model', 'J60')))

def calc_Segment_Revenue_Model_M61(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'M58:M60'))

def calc_Segment_Revenue_Model_N61(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'N58:N60'))

def calc_Segment_Revenue_Model_O61(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'O58:O60'))

def calc_Segment_Revenue_Model_Q61(ctx):
    return xl_sub(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'Q58')), xl_ref(ctx.cell('Segment Revenue Model', 'Q59'))), xl_ref(ctx.cell('Segment Revenue Model', 'Q60')))

def calc_Segment_Revenue_Model_R61(ctx):
    return xl_sub(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'R58')), xl_ref(ctx.cell('Segment Revenue Model', 'R59'))), xl_ref(ctx.cell('Segment Revenue Model', 'R60')))

def calc_Segment_Revenue_Model_S61(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'S58:S60'))

def calc_Segment_Revenue_Model_T61(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'T58:T60'))

def calc_Segment_Revenue_Model_V61(ctx):
    return xl_add(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'V58')), xl_ref(ctx.cell('Segment Revenue Model', 'V59'))), xl_ref(ctx.cell('Segment Revenue Model', 'V60')))

def calc_Segment_Revenue_Model_W61(ctx):
    return xl_add(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'W58')), xl_ref(ctx.cell('Segment Revenue Model', 'W59'))), xl_ref(ctx.cell('Segment Revenue Model', 'W60')))

def calc_Segment_Revenue_Model_X61(ctx):
    return xl_add(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'X58')), xl_ref(ctx.cell('Segment Revenue Model', 'X59'))), xl_ref(ctx.cell('Segment Revenue Model', 'X60')))

def calc_Segment_Revenue_Model_Y61(ctx):
    return xl_add(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'Y58')), xl_ref(ctx.cell('Segment Revenue Model', 'Y59'))), xl_ref(ctx.cell('Segment Revenue Model', 'Y60')))

def calc_Segment_Revenue_Model_F72(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'F65')), xl_ref(ctx.cell('Segment Revenue Model', 'F66'))), xl_ref(ctx.cell('Segment Revenue Model', 'F67'))), xl_ref(ctx.cell('Segment Revenue Model', 'F68'))), xl_ref(ctx.cell('Segment Revenue Model', 'F69'))), xl_ref(ctx.cell('Segment Revenue Model', 'F70'))), xl_ref(ctx.cell('Segment Revenue Model', 'F71')))

def calc_Segment_Revenue_Model_K72(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'K65')), xl_ref(ctx.cell('Segment Revenue Model', 'K66'))), xl_ref(ctx.cell('Segment Revenue Model', 'K67'))), xl_ref(ctx.cell('Segment Revenue Model', 'K68'))), xl_ref(ctx.cell('Segment Revenue Model', 'K69'))), xl_ref(ctx.cell('Segment Revenue Model', 'K70'))), xl_ref(ctx.cell('Segment Revenue Model', 'K71')))

def calc_Segment_Revenue_Model_P72(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'P65')), xl_ref(ctx.cell('Segment Revenue Model', 'P66'))), xl_ref(ctx.cell('Segment Revenue Model', 'P67'))), xl_ref(ctx.cell('Segment Revenue Model', 'P68'))), xl_ref(ctx.cell('Segment Revenue Model', 'P69'))), xl_ref(ctx.cell('Segment Revenue Model', 'P70'))), xl_ref(ctx.cell('Segment Revenue Model', 'P71')))

def calc_Segment_Revenue_Model_B74(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'B72:B73'))

def calc_Segment_Revenue_Model_C74(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'C72:C73'))

def calc_Segment_Revenue_Model_D74(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'D72:D73'))

def calc_Segment_Revenue_Model_E74(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'E72:E73'))

def calc_Segment_Revenue_Model_G74(ctx):
    return xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'G72')), xl_ref(ctx.cell('Segment Revenue Model', 'G73')))

def calc_Segment_Revenue_Model_H74(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'H72:H73'))

def calc_Segment_Revenue_Model_I74(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'I72:I73'))

def calc_Segment_Revenue_Model_J74(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'J72:J73'))

def calc_Segment_Revenue_Model_L74(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'L72:L73'))

def calc_Segment_Revenue_Model_M74(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'M72:M73'))

def calc_Segment_Revenue_Model_N74(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'N72:N73'))

def calc_Segment_Revenue_Model_O74(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'O72:O73'))

def calc_Segment_Revenue_Model_Q74(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'Q72:Q73'))

def calc_Segment_Revenue_Model_R74(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'R72:R73'))

def calc_Segment_Revenue_Model_S74(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'S72:S73'))

def calc_Segment_Revenue_Model_U72(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'Q72:T72'))

def calc_Segment_Revenue_Model_T74(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'T72:T73'))

def calc_INCOME_STATEMENT_F6(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'D6')), xl_ref(ctx.cell('INCOME STATEMENT', 'E6')))

def calc_INCOME_STATEMENT_D7(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'D6')), xl_ref(ctx.cell('INCOME STATEMENT', 'H6'))), 100)

def calc_INCOME_STATEMENT_K8(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'L6')), xl_ref(ctx.cell('INCOME STATEMENT', 'D6'))), 1), 100)

def calc_INCOME_STATEMENT_M6(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'K6')), xl_ref(ctx.cell('INCOME STATEMENT', 'L6')))

def calc_INCOME_STATEMENT_T6(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'R6')), xl_ref(ctx.cell('INCOME STATEMENT', 'S6')))

def calc_INCOME_STATEMENT_R7(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'R6')), xl_ref(ctx.cell('INCOME STATEMENT', 'V6'))), 100)

def calc_INCOME_STATEMENT_R8(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'R6')), xl_ref(ctx.cell('INCOME STATEMENT', 'L6'))), 1), 100)

def calc_INCOME_STATEMENT_AA6(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'Y6')), xl_ref(ctx.cell('INCOME STATEMENT', 'Z6')))

def calc_INCOME_STATEMENT_Y8(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Y6')), xl_ref(ctx.cell('INCOME STATEMENT', 'R6'))), 1), 100)

def calc_Ratio_Analysis_I21(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AF6')), xl_ref(ctx.cell('INCOME STATEMENT', 'Y6'))), 1)

def calc_INCOME_STATEMENT_B21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'B6')), xl_ref(ctx.cell('INCOME STATEMENT', 'B10')))

def calc_INCOME_STATEMENT_C21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'C6')), xl_ref(ctx.cell('INCOME STATEMENT', 'C10')))

def calc_INCOME_STATEMENT_E21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'E6')), xl_ref(ctx.cell('INCOME STATEMENT', 'E10')))

def calc_INCOME_STATEMENT_G21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'G6')), xl_ref(ctx.cell('INCOME STATEMENT', 'G10')))

def calc_INCOME_STATEMENT_I21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'I6')), xl_ref(ctx.cell('INCOME STATEMENT', 'I10')))

def calc_INCOME_STATEMENT_J21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'J6')), xl_ref(ctx.cell('INCOME STATEMENT', 'J10')))

def calc_INCOME_STATEMENT_L21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'L6')), xl_ref(ctx.cell('INCOME STATEMENT', 'L10')))

def calc_INCOME_STATEMENT_N21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'N6')), xl_ref(ctx.cell('INCOME STATEMENT', 'N10')))

def calc_INCOME_STATEMENT_P21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'P6')), xl_ref(ctx.cell('INCOME STATEMENT', 'P10')))

def calc_INCOME_STATEMENT_Q21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'Q6')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q10')))

def calc_INCOME_STATEMENT_S21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'S6')), xl_ref(ctx.cell('INCOME STATEMENT', 'S10')))

def calc_INCOME_STATEMENT_U21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'U6')), xl_ref(ctx.cell('INCOME STATEMENT', 'U10')))

def calc_INCOME_STATEMENT_W21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'W6')), xl_ref(ctx.cell('INCOME STATEMENT', 'W10')))

def calc_INCOME_STATEMENT_X21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'X6')), xl_ref(ctx.cell('INCOME STATEMENT', 'X10')))

def calc_INCOME_STATEMENT_Z21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'Z6')), xl_ref(ctx.cell('INCOME STATEMENT', 'Z10')))

def calc_INCOME_STATEMENT_AB21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AB6')), xl_ref(ctx.cell('INCOME STATEMENT', 'AB10')))

def calc_INCOME_STATEMENT_F11(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'D11')), xl_ref(ctx.cell('INCOME STATEMENT', 'E11')))

def calc_INCOME_STATEMENT_D12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'D11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'D6'))))

def calc_INCOME_STATEMENT_M11(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'K11')), xl_ref(ctx.cell('INCOME STATEMENT', 'L11')))

def calc_INCOME_STATEMENT_K12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'K11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'L6'))))

def calc_INCOME_STATEMENT_T11(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'R11')), xl_ref(ctx.cell('INCOME STATEMENT', 'S11')))

def calc_INCOME_STATEMENT_R12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'R11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'R6'))))

def calc_INCOME_STATEMENT_AA11(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'Y11')), xl_ref(ctx.cell('INCOME STATEMENT', 'Z11')))

def calc_INCOME_STATEMENT_Y12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Y11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'Y6'))))

def calc_INCOME_STATEMENT_F13(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'E13')), xl_ref(ctx.cell('INCOME STATEMENT', 'D13')))

def calc_INCOME_STATEMENT_D14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'D13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'D6'))))

def calc_INCOME_STATEMENT_M13(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'L13')), xl_ref(ctx.cell('INCOME STATEMENT', 'K13')))

def calc_INCOME_STATEMENT_K14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'K13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'L6'))))

def calc_INCOME_STATEMENT_T13(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'S13')), xl_ref(ctx.cell('INCOME STATEMENT', 'R13')))

def calc_INCOME_STATEMENT_R14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'R13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'R6'))))

def calc_INCOME_STATEMENT_AA13(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'Z13')), xl_ref(ctx.cell('INCOME STATEMENT', 'Y13')))

def calc_INCOME_STATEMENT_Y14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Y13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'Y6'))))

def calc_INCOME_STATEMENT_F15(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'E15')), xl_ref(ctx.cell('INCOME STATEMENT', 'D15')))

def calc_INCOME_STATEMENT_D16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'D15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'D6'))))

def calc_INCOME_STATEMENT_M15(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'L15')), xl_ref(ctx.cell('INCOME STATEMENT', 'K15')))

def calc_INCOME_STATEMENT_K16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'K15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'L6'))))

def calc_INCOME_STATEMENT_T15(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'S15')), xl_ref(ctx.cell('INCOME STATEMENT', 'R15')))

def calc_INCOME_STATEMENT_R16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'R15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'R6'))))

def calc_INCOME_STATEMENT_AA15(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'Z15')), xl_ref(ctx.cell('INCOME STATEMENT', 'Y15')))

def calc_INCOME_STATEMENT_Y16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Y15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'Y6'))))

def calc_INCOME_STATEMENT_F17(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'E17')), xl_ref(ctx.cell('INCOME STATEMENT', 'D17')))

def calc_INCOME_STATEMENT_D18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'D17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'D6'))))

def calc_INCOME_STATEMENT_M17(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'L17')), xl_ref(ctx.cell('INCOME STATEMENT', 'K17')))

def calc_INCOME_STATEMENT_K18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'K17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'L6'))))

def calc_INCOME_STATEMENT_T17(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'S17')), xl_ref(ctx.cell('INCOME STATEMENT', 'R17')))

def calc_INCOME_STATEMENT_R18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'R17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'R6'))))

def calc_INCOME_STATEMENT_AA17(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'Z17')), xl_ref(ctx.cell('INCOME STATEMENT', 'Y17')))

def calc_INCOME_STATEMENT_Y18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Y17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'Y6'))))

def calc_INCOME_STATEMENT_D10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'D11')), xl_ref(ctx.cell('INCOME STATEMENT', 'D13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'D15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'D17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'D19')))

def calc_INCOME_STATEMENT_F19(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'E19')), xl_ref(ctx.cell('INCOME STATEMENT', 'D19')))

def calc_INCOME_STATEMENT_D20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'D19')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'D6'))))

def calc_INCOME_STATEMENT_M19(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'L19')), xl_ref(ctx.cell('INCOME STATEMENT', 'K19')))

def calc_INCOME_STATEMENT_K20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'K19')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'L6'))))

def calc_INCOME_STATEMENT_R10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'R11')), xl_ref(ctx.cell('INCOME STATEMENT', 'R13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'R17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'R15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'R19')))

def calc_INCOME_STATEMENT_T19(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'S19')), xl_ref(ctx.cell('INCOME STATEMENT', 'R19')))

def calc_INCOME_STATEMENT_R20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'R19')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'R6'))))

def calc_INCOME_STATEMENT_Y10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'Y11')), xl_ref(ctx.cell('INCOME STATEMENT', 'Y13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'Y17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'Y15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'Y19')))

def calc_INCOME_STATEMENT_AA19(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'Z19')), xl_ref(ctx.cell('INCOME STATEMENT', 'Y19')))

def calc_INCOME_STATEMENT_Y20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Y19')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'Y6'))))

def calc_INCOME_STATEMENT_AD10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AD11')), xl_ref(ctx.cell('INCOME STATEMENT', 'AD13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AD15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AD17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AD19')))

def calc_INCOME_STATEMENT_AE10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AE11')), xl_ref(ctx.cell('INCOME STATEMENT', 'AE13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AE15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AE17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AE19')))

def calc_INCOME_STATEMENT_AF10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AF11')), xl_ref(ctx.cell('INCOME STATEMENT', 'AF13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AF15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AF17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AF19')))

def calc_INCOME_STATEMENT_AG10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AG11')), xl_ref(ctx.cell('INCOME STATEMENT', 'AG13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AG15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AG17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AG19')))

def calc_INCOME_STATEMENT_K25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'K21')), xl_ref(ctx.cell('INCOME STATEMENT', 'L6'))), 100)

def calc_INCOME_STATEMENT_F26(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'D26')), xl_ref(ctx.cell('INCOME STATEMENT', 'E26')))

def calc_INCOME_STATEMENT_M26(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'K26')), xl_ref(ctx.cell('INCOME STATEMENT', 'L26')))

def calc_INCOME_STATEMENT_T26(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'R26')), xl_ref(ctx.cell('INCOME STATEMENT', 'S26')))

def calc_INCOME_STATEMENT_AA26(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'Y26')), xl_ref(ctx.cell('INCOME STATEMENT', 'Z26')))

def calc_INCOME_STATEMENT_F27(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'D27')), xl_ref(ctx.cell('INCOME STATEMENT', 'E27')))

def calc_INCOME_STATEMENT_M27(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'K27')), xl_ref(ctx.cell('INCOME STATEMENT', 'L27')))

def calc_INCOME_STATEMENT_K28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'K21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'K26')), xl_ref(ctx.cell('INCOME STATEMENT', 'K27'))))

def calc_INCOME_STATEMENT_T27(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'R27')), xl_ref(ctx.cell('INCOME STATEMENT', 'S27')))

def calc_INCOME_STATEMENT_AA27(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'Y27')), xl_ref(ctx.cell('INCOME STATEMENT', 'Z27')))

def calc_INCOME_STATEMENT_F29(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'D29')), xl_ref(ctx.cell('INCOME STATEMENT', 'E29')))

def calc_INCOME_STATEMENT_M29(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'K29')), xl_ref(ctx.cell('INCOME STATEMENT', 'L29')))

def calc_INCOME_STATEMENT_T29(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'R29')), xl_ref(ctx.cell('INCOME STATEMENT', 'S29')))

def calc_INCOME_STATEMENT_AA29(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'Y29')), xl_ref(ctx.cell('INCOME STATEMENT', 'Z29')))

def calc_INCOME_STATEMENT_T31(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'R31')), xl_ref(ctx.cell('INCOME STATEMENT', 'S31')))

def calc_INCOME_STATEMENT_AA31(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'Y31')), xl_ref(ctx.cell('INCOME STATEMENT', 'Z31')))

def calc_INCOME_STATEMENT_F33(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'D33')), xl_ref(ctx.cell('INCOME STATEMENT', 'E33')))

def calc_INCOME_STATEMENT_M33(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'K33')), xl_ref(ctx.cell('INCOME STATEMENT', 'L33')))

def calc_INCOME_STATEMENT_T33(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'R33')), xl_ref(ctx.cell('INCOME STATEMENT', 'S33')))

def calc_INCOME_STATEMENT_AA33(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'Y33')), xl_ref(ctx.cell('INCOME STATEMENT', 'Z33')))

def calc_INCOME_STATEMENT_F34(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'D34')), xl_ref(ctx.cell('INCOME STATEMENT', 'E34')))

def calc_INCOME_STATEMENT_M34(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'K34')), xl_ref(ctx.cell('INCOME STATEMENT', 'L34')))

def calc_INCOME_STATEMENT_T34(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'R34')), xl_ref(ctx.cell('INCOME STATEMENT', 'S34')))

def calc_INCOME_STATEMENT_AA34(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'Y34')), xl_ref(ctx.cell('INCOME STATEMENT', 'Z34')))

def calc_INCOME_STATEMENT_F35(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'D35')), xl_ref(ctx.cell('INCOME STATEMENT', 'E35')))

def calc_INCOME_STATEMENT_M35(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'K35')), xl_ref(ctx.cell('INCOME STATEMENT', 'L35')))

def calc_INCOME_STATEMENT_T35(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'R35')), xl_ref(ctx.cell('INCOME STATEMENT', 'S35')))

def calc_INCOME_STATEMENT_AA35(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'Y35')), xl_ref(ctx.cell('INCOME STATEMENT', 'Z35')))

def calc_INCOME_STATEMENT_D47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'C47'))

def calc_INCOME_STATEMENT_C48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'C47')), 10)

def calc_INCOME_STATEMENT_U47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'T47'))

def calc_INCOME_STATEMENT_T48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'T47')), 10)

def calc_CASH_FOW_STATEMENT_I10(ctx):
    return xl_sub(xl_ref(ctx.cell('CASH FOW STATEMENT', 'I8')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'I9')))

def calc_Valuation_B27(ctx):
    return xl_ref(ctx.cell('CASH FOW STATEMENT', 'D13'))

def calc_Valuation_C27(ctx):
    return xl_ref(ctx.cell('CASH FOW STATEMENT', 'E13'))

def calc_Valuation_D27(ctx):
    return xl_ref(ctx.cell('CASH FOW STATEMENT', 'F13'))

def calc_Valuation_E27(ctx):
    return xl_ref(ctx.cell('CASH FOW STATEMENT', 'G13'))

def calc_Valuation_F27(ctx):
    return xl_ref(ctx.cell('CASH FOW STATEMENT', 'H13'))

def calc_Valuation_G27(ctx):
    return xl_ref(ctx.cell('CASH FOW STATEMENT', 'I13'))

def calc_Valuation_D25(ctx):
    return xl_ref(ctx.cell('CASH FOW STATEMENT', 'F20'))

def calc_Valuation_E25(ctx):
    return xl_ref(ctx.cell('CASH FOW STATEMENT', 'G20'))

def calc_Valuation_F25(ctx):
    return xl_ref(ctx.cell('CASH FOW STATEMENT', 'H20'))

def calc_Valuation_G25(ctx):
    return xl_ref(ctx.cell('CASH FOW STATEMENT', 'I20'))

def calc_BALANCESHEET_B8(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'B9:B10'))

def calc_BALANCESHEET_C8(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'C9:C10'))

def calc_BALANCESHEET_D8(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'D9:D10'))

def calc_BALANCESHEET_E8(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'E9:E10'))

def calc_BALANCESHEET_F8(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'F9:F10'))

def calc_BALANCESHEET_G8(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'G9:G10'))

def calc_BALANCESHEET_H8(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'H9:H10'))

def calc_BALANCESHEET_I8(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'I9:I10'))

def calc_BALANCESHEET_J8(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'J9:J10'))

def calc_BALANCESHEET_B24(ctx):
    return xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'B9')), xl_ref(ctx.cell('BALANCESHEET', 'B10'))), xl_ref(ctx.cell('BALANCESHEET', 'B19'))), xl_ref(ctx.cell('BALANCESHEET', 'B23')))

def calc_CASH_FOW_STATEMENT_B21(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'C19')), xl_ref(ctx.cell('BALANCESHEET', 'B19')))

def calc_BALANCESHEET_C24(ctx):
    return xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'C9')), xl_ref(ctx.cell('BALANCESHEET', 'C10'))), xl_ref(ctx.cell('BALANCESHEET', 'C19'))), xl_ref(ctx.cell('BALANCESHEET', 'C23')))

def calc_CASH_FOW_STATEMENT_C21(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'D19')), xl_ref(ctx.cell('BALANCESHEET', 'C19')))

def calc_BALANCESHEET_D24(ctx):
    return xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'D9')), xl_ref(ctx.cell('BALANCESHEET', 'D10'))), xl_ref(ctx.cell('BALANCESHEET', 'D19'))), xl_ref(ctx.cell('BALANCESHEET', 'D23')))

def calc_CASH_FOW_STATEMENT_D21(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'E19')), xl_ref(ctx.cell('BALANCESHEET', 'D19')))

def calc_BALANCESHEET_E24(ctx):
    return xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'E9')), xl_ref(ctx.cell('BALANCESHEET', 'E10'))), xl_ref(ctx.cell('BALANCESHEET', 'E19'))), xl_ref(ctx.cell('BALANCESHEET', 'E23')))

def calc_Valuation_G10(ctx):
    return xl_ref(ctx.cell('BALANCESHEET', 'E19'))

def calc_CASH_FOW_STATEMENT_E21(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'F19')), xl_ref(ctx.cell('BALANCESHEET', 'E19')))

def calc_BALANCESHEET_F24(ctx):
    return xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'F9')), xl_ref(ctx.cell('BALANCESHEET', 'F10'))), xl_ref(ctx.cell('BALANCESHEET', 'F19'))), xl_ref(ctx.cell('BALANCESHEET', 'F23')))

def calc_CASH_FOW_STATEMENT_F21(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'G19')), xl_ref(ctx.cell('BALANCESHEET', 'F19')))

def calc_BALANCESHEET_G24(ctx):
    return xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'G9')), xl_ref(ctx.cell('BALANCESHEET', 'G10'))), xl_ref(ctx.cell('BALANCESHEET', 'G19'))), xl_ref(ctx.cell('BALANCESHEET', 'G23')))

def calc_Ratio_Analysis_G15(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AD26')), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'G19')), xl_ref(ctx.cell('BALANCESHEET', 'F19'))))

def calc_CASH_FOW_STATEMENT_G21(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'H19')), xl_ref(ctx.cell('BALANCESHEET', 'G19')))

def calc_BALANCESHEET_H24(ctx):
    return xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'H9')), xl_ref(ctx.cell('BALANCESHEET', 'H10'))), xl_ref(ctx.cell('BALANCESHEET', 'H19'))), xl_ref(ctx.cell('BALANCESHEET', 'H23')))

def calc_Ratio_Analysis_H15(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AE26')), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'H19')), xl_ref(ctx.cell('BALANCESHEET', 'G19'))))

def calc_CASH_FOW_STATEMENT_H21(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'I19')), xl_ref(ctx.cell('BALANCESHEET', 'H19')))

def calc_BALANCESHEET_I24(ctx):
    return xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'I9')), xl_ref(ctx.cell('BALANCESHEET', 'I10'))), xl_ref(ctx.cell('BALANCESHEET', 'I19'))), xl_ref(ctx.cell('BALANCESHEET', 'I23')))

def calc_Ratio_Analysis_I15(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AF26')), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'I19')), xl_ref(ctx.cell('BALANCESHEET', 'H19'))))

def calc_CASH_FOW_STATEMENT_I21(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'J19')), xl_ref(ctx.cell('BALANCESHEET', 'I19')))

def calc_BALANCESHEET_J24(ctx):
    return xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'J9')), xl_ref(ctx.cell('BALANCESHEET', 'J10'))), xl_ref(ctx.cell('BALANCESHEET', 'J19'))), xl_ref(ctx.cell('BALANCESHEET', 'J23')))

def calc_Ratio_Analysis_J15(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AG26')), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'J19')), xl_ref(ctx.cell('BALANCESHEET', 'I19'))))

def calc_BALANCESHEET_B27(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'B30:B31'))

def calc_BALANCESHEET_C27(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'C30:C31'))

def calc_BALANCESHEET_D27(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'D30:D31'))

def calc_BALANCESHEET_E27(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'E30:E31'))

def calc_BALANCESHEET_F27(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'F30:F31'))

def calc_BALANCESHEET_G27(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'G30:G31'))

def calc_BALANCESHEET_H27(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'H30:H31'))

def calc_BALANCESHEET_I27(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'I30:I31'))

def calc_BALANCESHEET_J27(ctx):
    return xl_sum(ctx.range('BALANCESHEET', 'J30:J31'))

def calc_Ratio_Analysis_F9(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'F36')), xl_ref(ctx.cell('BALANCESHEET', 'F43')))

def calc_Ratio_Analysis_G9(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'G36')), xl_ref(ctx.cell('BALANCESHEET', 'G43')))

def calc_Ratio_Analysis_H9(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'H36')), xl_ref(ctx.cell('BALANCESHEET', 'H43')))

def calc_Ratio_Analysis_I9(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'I36')), xl_ref(ctx.cell('BALANCESHEET', 'I43')))

def calc_Ratio_Analysis_J9(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'J36')), xl_ref(ctx.cell('BALANCESHEET', 'J43')))

def calc_BALANCESHEET_B43(ctx):
    return xl_add(xl_ref(ctx.cell('BALANCESHEET', 'B44')), xl_ref(ctx.cell('BALANCESHEET', 'B52')))

def calc_BALANCESHEET_B58(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'B36')), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'B44')), xl_ref(ctx.cell('BALANCESHEET', 'B52'))))

def calc_BALANCESHEET_C43(ctx):
    return xl_add(xl_ref(ctx.cell('BALANCESHEET', 'C44')), xl_ref(ctx.cell('BALANCESHEET', 'C52')))

def calc_BALANCESHEET_C58(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'C36')), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'C44')), xl_ref(ctx.cell('BALANCESHEET', 'C52'))))

def calc_BALANCESHEET_D43(ctx):
    return xl_add(xl_ref(ctx.cell('BALANCESHEET', 'D44')), xl_ref(ctx.cell('BALANCESHEET', 'D52')))

def calc_BALANCESHEET_D58(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'D36')), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'D44')), xl_ref(ctx.cell('BALANCESHEET', 'D52'))))

def calc_BALANCESHEET_E43(ctx):
    return xl_add(xl_ref(ctx.cell('BALANCESHEET', 'E44')), xl_ref(ctx.cell('BALANCESHEET', 'E52')))

def calc_BALANCESHEET_E58(ctx):
    return xl_sub(xl_ref(ctx.cell('BALANCESHEET', 'E36')), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'E44')), xl_ref(ctx.cell('BALANCESHEET', 'E52'))))

def calc_Debt_Schedule_B21(ctx):
    return xl_add(xl_ref(ctx.cell('Debt Schedule', 'B20')), xl_ref(ctx.cell('Debt Schedule', 'B11')))

def calc_Debt_Schedule_C21(ctx):
    return xl_add(xl_ref(ctx.cell('Debt Schedule', 'C20')), xl_ref(ctx.cell('Debt Schedule', 'C11')))

def calc_Debt_Schedule_D21(ctx):
    return xl_add(xl_ref(ctx.cell('Debt Schedule', 'D20')), xl_ref(ctx.cell('Debt Schedule', 'D11')))

def calc_Debt_Schedule_E21(ctx):
    return xl_add(xl_ref(ctx.cell('Debt Schedule', 'E20')), xl_ref(ctx.cell('Debt Schedule', 'E11')))

def calc_Debt_Schedule_F21(ctx):
    return xl_add(xl_ref(ctx.cell('Debt Schedule', 'F20')), xl_ref(ctx.cell('Debt Schedule', 'F11')))

def calc_Debt_Schedule_G21(ctx):
    return xl_add(xl_ref(ctx.cell('Debt Schedule', 'G20')), xl_ref(ctx.cell('Debt Schedule', 'G11')))

def calc_Debt_Schedule_H21(ctx):
    return xl_add(xl_ref(ctx.cell('Debt Schedule', 'H20')), xl_ref(ctx.cell('Debt Schedule', 'H11')))

def calc_Debt_Schedule_I21(ctx):
    return xl_add(xl_ref(ctx.cell('Debt Schedule', 'I20')), xl_ref(ctx.cell('Debt Schedule', 'I11')))

def calc_Valuation_G54(ctx):
    return xl_uminus(xl_ref(ctx.cell('Valuation', 'G11')))

def calc_Valuation_E37(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'E19')), xl_ref(ctx.cell('Valuation', 'D19'))), 1)

def calc_Valuation_D40(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'D24')), xl_ref(ctx.cell('Valuation', 'C24'))), 1)

def calc_Valuation_D45(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'D24')), xl_ref(ctx.cell('Valuation', 'D19')))

def calc_Valuation_E40(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'E24')), xl_ref(ctx.cell('Valuation', 'D24'))), 1)

def calc_Valuation_E45(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'E24')), xl_ref(ctx.cell('Valuation', 'E19')))

def calc_Valuation_F40(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'F24')), xl_ref(ctx.cell('Valuation', 'E24'))), 1)

def calc_Valuation_G40(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'G24')), xl_ref(ctx.cell('Valuation', 'F24'))), 1)

def calc_PRESENTATION_G65(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'G14'))

def calc_PRESENTATION_H65(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'H14'))

def calc_PRESENTATION_I65(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'I14'))

def calc_PRESENTATION_J65(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'J14'))

def calc_PRESENTATION_I58(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'I23'))

def calc_PRESENTATION_J58(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'J23'))

def calc_PRESENTATION_K17(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'K15')), xl_ref(ctx.cell('PRESENTATION', 'K16')))

def calc_PRESENTATION_P17(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'P15')), xl_ref(ctx.cell('PRESENTATION', 'P16')))

def calc_PRESENTATION_U15(ctx):
    return xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'Q15')), xl_ref(ctx.cell('PRESENTATION', 'R15'))), xl_ref(ctx.cell('PRESENTATION', 'S15'))), xl_ref(ctx.cell('PRESENTATION', 'T15')))

def calc_PRESENTATION_T17(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'T15')), xl_ref(ctx.cell('PRESENTATION', 'T16')))

def calc_PRESENTATION_F17(ctx):
    return xl_sum(ctx.range('PRESENTATION', 'B17:E17'))

def calc_PRESENTATION_O25(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'M25')), xl_ref(ctx.cell('PRESENTATION', 'N25')))

def calc_PRESENTATION_AC25(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'AA25')), xl_ref(ctx.cell('PRESENTATION', 'AB25')))

def calc_PRESENTATION_B33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'B32')), xl_ref(ctx.cell('PRESENTATION', 'B25'))), 100)

def calc_PRESENTATION_B36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'B32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'B34')), xl_ref(ctx.cell('PRESENTATION', 'B35'))))

def calc_PRESENTATION_C33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'C32')), xl_ref(ctx.cell('PRESENTATION', 'C25'))), 100)

def calc_PRESENTATION_C36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'C32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'C34')), xl_ref(ctx.cell('PRESENTATION', 'C35'))))

def calc_PRESENTATION_E33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'E32')), xl_ref(ctx.cell('PRESENTATION', 'E25'))), 100)

def calc_PRESENTATION_E36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'E32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'E34')), xl_ref(ctx.cell('PRESENTATION', 'E35'))))

def calc_PRESENTATION_G33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'G32')), xl_ref(ctx.cell('PRESENTATION', 'G25'))), 100)

def calc_PRESENTATION_G36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'G32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'G34')), xl_ref(ctx.cell('PRESENTATION', 'G35'))))

def calc_PRESENTATION_I33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'I32')), xl_ref(ctx.cell('PRESENTATION', 'I25'))), 100)

def calc_PRESENTATION_I36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'I32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'I34')), xl_ref(ctx.cell('PRESENTATION', 'I35'))))

def calc_PRESENTATION_J33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'J32')), xl_ref(ctx.cell('PRESENTATION', 'J25'))), 100)

def calc_PRESENTATION_J36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'J32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'J34')), xl_ref(ctx.cell('PRESENTATION', 'J35'))))

def calc_PRESENTATION_L33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'L32')), xl_ref(ctx.cell('PRESENTATION', 'L25'))), 100)

def calc_PRESENTATION_L36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'L32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'L34')), xl_ref(ctx.cell('PRESENTATION', 'L35'))))

def calc_PRESENTATION_N33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'N32')), xl_ref(ctx.cell('PRESENTATION', 'N25'))), 100)

def calc_PRESENTATION_N36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'N32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'N34')), xl_ref(ctx.cell('PRESENTATION', 'N35'))))

def calc_PRESENTATION_P33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'P32')), xl_ref(ctx.cell('PRESENTATION', 'P25'))), 100)

def calc_PRESENTATION_P36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'P32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'P34')), xl_ref(ctx.cell('PRESENTATION', 'P35'))))

def calc_PRESENTATION_Q33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'Q32')), xl_ref(ctx.cell('PRESENTATION', 'Q25'))), 100)

def calc_PRESENTATION_Q36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'Q32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'Q34')), xl_ref(ctx.cell('PRESENTATION', 'Q35'))))

def calc_PRESENTATION_S33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'S32')), xl_ref(ctx.cell('PRESENTATION', 'S25'))), 100)

def calc_PRESENTATION_S36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'S32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'S34')), xl_ref(ctx.cell('PRESENTATION', 'S35'))))

def calc_PRESENTATION_U33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'U32')), xl_ref(ctx.cell('PRESENTATION', 'U25'))), 100)

def calc_PRESENTATION_U36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'U32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'U34')), xl_ref(ctx.cell('PRESENTATION', 'U35'))))

def calc_PRESENTATION_W33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'W32')), xl_ref(ctx.cell('PRESENTATION', 'W25'))), 100)

def calc_PRESENTATION_W36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'W32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'W34')), xl_ref(ctx.cell('PRESENTATION', 'W35'))))

def calc_PRESENTATION_X33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'X32')), xl_ref(ctx.cell('PRESENTATION', 'X25'))), 100)

def calc_PRESENTATION_X36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'X32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'X34')), xl_ref(ctx.cell('PRESENTATION', 'X35'))))

def calc_PRESENTATION_Z33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'Z32')), xl_ref(ctx.cell('PRESENTATION', 'Z25'))), 100)

def calc_PRESENTATION_Z36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'Z32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'Z34')), xl_ref(ctx.cell('PRESENTATION', 'Z35'))))

def calc_PRESENTATION_AB33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'AB32')), xl_ref(ctx.cell('PRESENTATION', 'AB25'))), 100)

def calc_PRESENTATION_AB36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AB32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'AB34')), xl_ref(ctx.cell('PRESENTATION', 'AB35'))))

def calc_PRESENTATION_AD33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'AD32')), xl_ref(ctx.cell('PRESENTATION', 'AD25'))), 100)

def calc_PRESENTATION_AD36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AD32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'AD34')), xl_ref(ctx.cell('PRESENTATION', 'AD35'))))

def calc_PRESENTATION_AE33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'AE32')), xl_ref(ctx.cell('PRESENTATION', 'AE25'))), 100)

def calc_PRESENTATION_AE36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AE32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'AE34')), xl_ref(ctx.cell('PRESENTATION', 'AE35'))))

def calc_PRESENTATION_AF33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'AF32')), xl_ref(ctx.cell('PRESENTATION', 'AF25'))), 100)

def calc_PRESENTATION_AF36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AF32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'AF34')), xl_ref(ctx.cell('PRESENTATION', 'AF35'))))

def calc_PRESENTATION_AG33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'AG32')), xl_ref(ctx.cell('PRESENTATION', 'AG25'))), 100)

def calc_PRESENTATION_AG36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AG32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'AG34')), xl_ref(ctx.cell('PRESENTATION', 'AG35'))))

def calc_PRESENTATION_O27(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'M27')), xl_ref(ctx.cell('PRESENTATION', 'N27')))

def calc_PRESENTATION_AC27(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'AA27')), xl_ref(ctx.cell('PRESENTATION', 'AB27')))

def calc_PRESENTATION_AC28(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'AA28')), xl_ref(ctx.cell('PRESENTATION', 'AB28')))

def calc_PRESENTATION_H29(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'F29')), xl_ref(ctx.cell('PRESENTATION', 'G29')))

def calc_PRESENTATION_O29(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'M29')), xl_ref(ctx.cell('PRESENTATION', 'N29')))

def calc_PRESENTATION_V29(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'T29')), xl_ref(ctx.cell('PRESENTATION', 'U29')))

def calc_PRESENTATION_AC29(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'AA29')), xl_ref(ctx.cell('PRESENTATION', 'AB29')))

def calc_PRESENTATION_O30(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'M30')), xl_ref(ctx.cell('PRESENTATION', 'N30')))

def calc_PRESENTATION_AC30(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'AA30')), xl_ref(ctx.cell('PRESENTATION', 'AB30')))

def calc_PRESENTATION_D32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'D25')), xl_ref(ctx.cell('PRESENTATION', 'D26')))

def calc_PRESENTATION_F26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'F27')), xl_ref(ctx.cell('PRESENTATION', 'F28'))), xl_ref(ctx.cell('PRESENTATION', 'F29'))), xl_ref(ctx.cell('PRESENTATION', 'F30'))), xl_ref(ctx.cell('PRESENTATION', 'F31')))

def calc_PRESENTATION_H31(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'F31')), xl_ref(ctx.cell('PRESENTATION', 'G31')))

def calc_PRESENTATION_M26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'M27')), xl_ref(ctx.cell('PRESENTATION', 'M28'))), xl_ref(ctx.cell('PRESENTATION', 'M30'))), xl_ref(ctx.cell('PRESENTATION', 'M29'))), xl_ref(ctx.cell('PRESENTATION', 'M31')))

def calc_PRESENTATION_R32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'R25')), xl_ref(ctx.cell('PRESENTATION', 'R26')))

def calc_PRESENTATION_T26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'T27')), xl_ref(ctx.cell('PRESENTATION', 'T28'))), xl_ref(ctx.cell('PRESENTATION', 'T30'))), xl_ref(ctx.cell('PRESENTATION', 'T29'))), xl_ref(ctx.cell('PRESENTATION', 'T31')))

def calc_PRESENTATION_V31(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'T31')), xl_ref(ctx.cell('PRESENTATION', 'U31')))

def calc_PRESENTATION_Y32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'Y25')), xl_ref(ctx.cell('PRESENTATION', 'Y26')))

def calc_PRESENTATION_AA26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'AA27')), xl_ref(ctx.cell('PRESENTATION', 'AA28'))), xl_ref(ctx.cell('PRESENTATION', 'AA30'))), xl_ref(ctx.cell('PRESENTATION', 'AA29'))), xl_ref(ctx.cell('PRESENTATION', 'AA31')))

def calc_PRESENTATION_AC31(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'AA31')), xl_ref(ctx.cell('PRESENTATION', 'AB31')))

def calc_PRESENTATION_H34(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'F34')), xl_ref(ctx.cell('PRESENTATION', 'G34')))

def calc_PRESENTATION_O34(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'M34')), xl_ref(ctx.cell('PRESENTATION', 'N34')))

def calc_PRESENTATION_V34(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'T34')), xl_ref(ctx.cell('PRESENTATION', 'U34')))

def calc_PRESENTATION_AC34(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'AA34')), xl_ref(ctx.cell('PRESENTATION', 'AB34')))

def calc_PRESENTATION_H35(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'F35')), xl_ref(ctx.cell('PRESENTATION', 'G35')))

def calc_PRESENTATION_O35(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'M35')), xl_ref(ctx.cell('PRESENTATION', 'N35')))

def calc_PRESENTATION_K38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'K36')), xl_ref(ctx.cell('PRESENTATION', 'K37')))

def calc_PRESENTATION_V35(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'T35')), xl_ref(ctx.cell('PRESENTATION', 'U35')))

def calc_PRESENTATION_H37(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'F37')), xl_ref(ctx.cell('PRESENTATION', 'G37')))

def calc_PRESENTATION_O37(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'M37')), xl_ref(ctx.cell('PRESENTATION', 'N37')))

def calc_PRESENTATION_AC37(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'AA37')), xl_ref(ctx.cell('PRESENTATION', 'AB37')))

def calc_PRESENTATION_AC41(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'AA41')), xl_ref(ctx.cell('PRESENTATION', 'AB41')))

def calc_PRESENTATION_H42(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'F42')), xl_ref(ctx.cell('PRESENTATION', 'G42')))

def calc_PRESENTATION_AC42(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'AA42')), xl_ref(ctx.cell('PRESENTATION', 'AB42')))

def calc_PRESENTATION_H43(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'F43')), xl_ref(ctx.cell('PRESENTATION', 'G43')))

def calc_Segment_Revenue_Model_T8(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'U8')), xl_add(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'Q8')), xl_ref(ctx.cell('Segment Revenue Model', 'R8'))), xl_ref(ctx.cell('Segment Revenue Model', 'S8'))))

def calc_Segment_Revenue_Model_V8(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'V24')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'U8')))

def calc_Segment_Revenue_Model_T9(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'U9')), xl_add(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'Q9')), xl_ref(ctx.cell('Segment Revenue Model', 'R9'))), xl_ref(ctx.cell('Segment Revenue Model', 'S9'))))

def calc_Segment_Revenue_Model_V9(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'V25')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'U9')))

def calc_Segment_Revenue_Model_T10(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'U10')), xl_add(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'Q10')), xl_ref(ctx.cell('Segment Revenue Model', 'R10'))), xl_ref(ctx.cell('Segment Revenue Model', 'S10'))))

def calc_Segment_Revenue_Model_V10(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'V26')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'U10')))

def calc_Segment_Revenue_Model_T11(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'U11')), xl_add(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'Q11')), xl_ref(ctx.cell('Segment Revenue Model', 'R11'))), xl_ref(ctx.cell('Segment Revenue Model', 'S11'))))

def calc_Segment_Revenue_Model_V11(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'V27')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'U11')))

def calc_Segment_Revenue_Model_T28(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'T12')), xl_ref(ctx.cell('Segment Revenue Model', 'O12'))), 1)

def calc_Segment_Revenue_Model_W12(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'W28')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'V12')))

def calc_Segment_Revenue_Model_T13(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'U13')), xl_add(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'Q13')), xl_ref(ctx.cell('Segment Revenue Model', 'R13'))), xl_ref(ctx.cell('Segment Revenue Model', 'S13'))))

def calc_Segment_Revenue_Model_V13(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'V29')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'U13')))

def calc_Segment_Revenue_Model_K17(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'K15')), xl_ref(ctx.cell('Segment Revenue Model', 'K16')))

def calc_Segment_Revenue_Model_K31(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'K15')), xl_ref(ctx.cell('Segment Revenue Model', 'F15'))), 1)

def calc_Segment_Revenue_Model_T14(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'U14')), xl_add(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'Q14')), xl_ref(ctx.cell('Segment Revenue Model', 'R14'))), xl_ref(ctx.cell('Segment Revenue Model', 'S14'))))

def calc_Segment_Revenue_Model_V14(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'V30')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'U14')))

def calc_Segment_Revenue_Model_P17(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'P15')), xl_ref(ctx.cell('Segment Revenue Model', 'P16')))

def calc_Segment_Revenue_Model_P31(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'P15')), xl_ref(ctx.cell('Segment Revenue Model', 'K15'))), 1)

def calc_Segment_Revenue_Model_F17(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'B17:E17'))

def calc_Segment_Revenue_Model_G33(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'G17')), xl_ref(ctx.cell('Segment Revenue Model', 'B17'))), 1)

def calc_Segment_Revenue_Model_G37(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'G8')), xl_ref(ctx.cell('Segment Revenue Model', 'G17')))

def calc_Segment_Revenue_Model_G38(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'G9')), xl_ref(ctx.cell('Segment Revenue Model', 'G17')))

def calc_Segment_Revenue_Model_G39(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'G10')), xl_ref(ctx.cell('Segment Revenue Model', 'G17')))

def calc_Segment_Revenue_Model_G40(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'G11')), xl_ref(ctx.cell('Segment Revenue Model', 'G17')))

def calc_Segment_Revenue_Model_G41(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'G12')), xl_ref(ctx.cell('Segment Revenue Model', 'G17')))

def calc_Segment_Revenue_Model_G42(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'G13')), xl_ref(ctx.cell('Segment Revenue Model', 'G17')))

def calc_Segment_Revenue_Model_G43(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'G14')), xl_ref(ctx.cell('Segment Revenue Model', 'G17')))

def calc_Segment_Revenue_Model_G44(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'G15')), xl_ref(ctx.cell('Segment Revenue Model', 'G17')))

def calc_Segment_Revenue_Model_G45(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'G16')), xl_ref(ctx.cell('Segment Revenue Model', 'G17')))

def calc_Segment_Revenue_Model_G46(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'G17')), xl_ref(ctx.cell('Segment Revenue Model', 'G17')))

def calc_Segment_Revenue_Model_H33(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'H17')), xl_ref(ctx.cell('Segment Revenue Model', 'C17'))), 1)

def calc_Segment_Revenue_Model_H37(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'H8')), xl_ref(ctx.cell('Segment Revenue Model', 'H17')))

def calc_Segment_Revenue_Model_H38(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'H9')), xl_ref(ctx.cell('Segment Revenue Model', 'H17')))

def calc_Segment_Revenue_Model_H39(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'H10')), xl_ref(ctx.cell('Segment Revenue Model', 'H17')))

def calc_Segment_Revenue_Model_H40(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'H11')), xl_ref(ctx.cell('Segment Revenue Model', 'H17')))

def calc_Segment_Revenue_Model_H41(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'H12')), xl_ref(ctx.cell('Segment Revenue Model', 'H17')))

def calc_Segment_Revenue_Model_H42(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'H13')), xl_ref(ctx.cell('Segment Revenue Model', 'H17')))

def calc_Segment_Revenue_Model_H43(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'H14')), xl_ref(ctx.cell('Segment Revenue Model', 'H17')))

def calc_Segment_Revenue_Model_H44(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'H15')), xl_ref(ctx.cell('Segment Revenue Model', 'H17')))

def calc_Segment_Revenue_Model_H45(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'H16')), xl_ref(ctx.cell('Segment Revenue Model', 'H17')))

def calc_Segment_Revenue_Model_H46(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'H17')), xl_ref(ctx.cell('Segment Revenue Model', 'H17')))

def calc_Segment_Revenue_Model_I33(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'I17')), xl_ref(ctx.cell('Segment Revenue Model', 'D17'))), 1)

def calc_Segment_Revenue_Model_I37(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'I8')), xl_ref(ctx.cell('Segment Revenue Model', 'I17')))

def calc_Segment_Revenue_Model_I38(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'I9')), xl_ref(ctx.cell('Segment Revenue Model', 'I17')))

def calc_Segment_Revenue_Model_I39(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'I10')), xl_ref(ctx.cell('Segment Revenue Model', 'I17')))

def calc_Segment_Revenue_Model_I40(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'I11')), xl_ref(ctx.cell('Segment Revenue Model', 'I17')))

def calc_Segment_Revenue_Model_I41(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'I12')), xl_ref(ctx.cell('Segment Revenue Model', 'I17')))

def calc_Segment_Revenue_Model_I42(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'I13')), xl_ref(ctx.cell('Segment Revenue Model', 'I17')))

def calc_Segment_Revenue_Model_I43(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'I14')), xl_ref(ctx.cell('Segment Revenue Model', 'I17')))

def calc_Segment_Revenue_Model_I44(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'I15')), xl_ref(ctx.cell('Segment Revenue Model', 'I17')))

def calc_Segment_Revenue_Model_I45(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'I16')), xl_ref(ctx.cell('Segment Revenue Model', 'I17')))

def calc_Segment_Revenue_Model_I46(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'I17')), xl_ref(ctx.cell('Segment Revenue Model', 'I17')))

def calc_Segment_Revenue_Model_J33(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'J17')), xl_ref(ctx.cell('Segment Revenue Model', 'E17'))), 1)

def calc_Segment_Revenue_Model_J37(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'J8')), xl_ref(ctx.cell('Segment Revenue Model', 'J17')))

def calc_Segment_Revenue_Model_J38(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'J9')), xl_ref(ctx.cell('Segment Revenue Model', 'J17')))

def calc_Segment_Revenue_Model_J39(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'J10')), xl_ref(ctx.cell('Segment Revenue Model', 'J17')))

def calc_Segment_Revenue_Model_J40(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'J11')), xl_ref(ctx.cell('Segment Revenue Model', 'J17')))

def calc_Segment_Revenue_Model_J41(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'J12')), xl_ref(ctx.cell('Segment Revenue Model', 'J17')))

def calc_Segment_Revenue_Model_J42(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'J13')), xl_ref(ctx.cell('Segment Revenue Model', 'J17')))

def calc_Segment_Revenue_Model_J43(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'J14')), xl_ref(ctx.cell('Segment Revenue Model', 'J17')))

def calc_Segment_Revenue_Model_J44(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'J15')), xl_ref(ctx.cell('Segment Revenue Model', 'J17')))

def calc_Segment_Revenue_Model_J45(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'J16')), xl_ref(ctx.cell('Segment Revenue Model', 'J17')))

def calc_Segment_Revenue_Model_J46(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'J17')), xl_ref(ctx.cell('Segment Revenue Model', 'J17')))

def calc_Segment_Revenue_Model_L33(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'L17')), xl_ref(ctx.cell('Segment Revenue Model', 'G17'))), 1)

def calc_Segment_Revenue_Model_L37(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'L8')), xl_ref(ctx.cell('Segment Revenue Model', 'L17')))

def calc_Segment_Revenue_Model_L38(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'L9')), xl_ref(ctx.cell('Segment Revenue Model', 'L17')))

def calc_Segment_Revenue_Model_L39(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'L10')), xl_ref(ctx.cell('Segment Revenue Model', 'L17')))

def calc_Segment_Revenue_Model_L40(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'L11')), xl_ref(ctx.cell('Segment Revenue Model', 'L17')))

def calc_Segment_Revenue_Model_L41(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'L12')), xl_ref(ctx.cell('Segment Revenue Model', 'L17')))

def calc_Segment_Revenue_Model_L42(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'L13')), xl_ref(ctx.cell('Segment Revenue Model', 'L17')))

def calc_Segment_Revenue_Model_L43(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'L14')), xl_ref(ctx.cell('Segment Revenue Model', 'L17')))

def calc_Segment_Revenue_Model_L44(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'L15')), xl_ref(ctx.cell('Segment Revenue Model', 'L17')))

def calc_Segment_Revenue_Model_L45(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'L16')), xl_ref(ctx.cell('Segment Revenue Model', 'L17')))

def calc_Segment_Revenue_Model_L46(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'L17')), xl_ref(ctx.cell('Segment Revenue Model', 'L17')))

def calc_Segment_Revenue_Model_M33(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'M17')), xl_ref(ctx.cell('Segment Revenue Model', 'H17'))), 1)

def calc_Segment_Revenue_Model_M37(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'M8')), xl_ref(ctx.cell('Segment Revenue Model', 'M17')))

def calc_Segment_Revenue_Model_M38(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'M9')), xl_ref(ctx.cell('Segment Revenue Model', 'M17')))

def calc_Segment_Revenue_Model_M39(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'M10')), xl_ref(ctx.cell('Segment Revenue Model', 'M17')))

def calc_Segment_Revenue_Model_M40(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'M11')), xl_ref(ctx.cell('Segment Revenue Model', 'M17')))

def calc_Segment_Revenue_Model_M41(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'M12')), xl_ref(ctx.cell('Segment Revenue Model', 'M17')))

def calc_Segment_Revenue_Model_M42(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'M13')), xl_ref(ctx.cell('Segment Revenue Model', 'M17')))

def calc_Segment_Revenue_Model_M43(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'M14')), xl_ref(ctx.cell('Segment Revenue Model', 'M17')))

def calc_Segment_Revenue_Model_M44(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'M15')), xl_ref(ctx.cell('Segment Revenue Model', 'M17')))

def calc_Segment_Revenue_Model_M45(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'M16')), xl_ref(ctx.cell('Segment Revenue Model', 'M17')))

def calc_Segment_Revenue_Model_M46(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'M17')), xl_ref(ctx.cell('Segment Revenue Model', 'M17')))

def calc_Segment_Revenue_Model_N33(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'N17')), xl_ref(ctx.cell('Segment Revenue Model', 'I17'))), 1)

def calc_Segment_Revenue_Model_N37(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'N8')), xl_ref(ctx.cell('Segment Revenue Model', 'N17')))

def calc_Segment_Revenue_Model_N38(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'N9')), xl_ref(ctx.cell('Segment Revenue Model', 'N17')))

def calc_Segment_Revenue_Model_N39(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'N10')), xl_ref(ctx.cell('Segment Revenue Model', 'N17')))

def calc_Segment_Revenue_Model_N40(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'N11')), xl_ref(ctx.cell('Segment Revenue Model', 'N17')))

def calc_Segment_Revenue_Model_N41(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'N12')), xl_ref(ctx.cell('Segment Revenue Model', 'N17')))

def calc_Segment_Revenue_Model_N42(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'N13')), xl_ref(ctx.cell('Segment Revenue Model', 'N17')))

def calc_Segment_Revenue_Model_N43(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'N14')), xl_ref(ctx.cell('Segment Revenue Model', 'N17')))

def calc_Segment_Revenue_Model_N44(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'N15')), xl_ref(ctx.cell('Segment Revenue Model', 'N17')))

def calc_Segment_Revenue_Model_N45(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'N16')), xl_ref(ctx.cell('Segment Revenue Model', 'N17')))

def calc_Segment_Revenue_Model_N46(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'N17')), xl_ref(ctx.cell('Segment Revenue Model', 'N17')))

def calc_Segment_Revenue_Model_O33(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'O17')), xl_ref(ctx.cell('Segment Revenue Model', 'J17'))), 1)

def calc_Segment_Revenue_Model_O37(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'O8')), xl_ref(ctx.cell('Segment Revenue Model', 'O17')))

def calc_Segment_Revenue_Model_O38(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'O9')), xl_ref(ctx.cell('Segment Revenue Model', 'O17')))

def calc_Segment_Revenue_Model_O39(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'O10')), xl_ref(ctx.cell('Segment Revenue Model', 'O17')))

def calc_Segment_Revenue_Model_O40(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'O11')), xl_ref(ctx.cell('Segment Revenue Model', 'O17')))

def calc_Segment_Revenue_Model_O41(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'O12')), xl_ref(ctx.cell('Segment Revenue Model', 'O17')))

def calc_Segment_Revenue_Model_O42(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'O13')), xl_ref(ctx.cell('Segment Revenue Model', 'O17')))

def calc_Segment_Revenue_Model_O43(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'O14')), xl_ref(ctx.cell('Segment Revenue Model', 'O17')))

def calc_Segment_Revenue_Model_O44(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'O15')), xl_ref(ctx.cell('Segment Revenue Model', 'O17')))

def calc_Segment_Revenue_Model_O45(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'O16')), xl_ref(ctx.cell('Segment Revenue Model', 'O17')))

def calc_Segment_Revenue_Model_O46(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'O17')), xl_ref(ctx.cell('Segment Revenue Model', 'O17')))

def calc_Segment_Revenue_Model_Q33(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Q17')), xl_ref(ctx.cell('Segment Revenue Model', 'L17'))), 1)

def calc_Segment_Revenue_Model_Q37(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Q8')), xl_ref(ctx.cell('Segment Revenue Model', 'Q17')))

def calc_Segment_Revenue_Model_Q38(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Q9')), xl_ref(ctx.cell('Segment Revenue Model', 'Q17')))

def calc_Segment_Revenue_Model_Q39(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Q10')), xl_ref(ctx.cell('Segment Revenue Model', 'Q17')))

def calc_Segment_Revenue_Model_Q40(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Q11')), xl_ref(ctx.cell('Segment Revenue Model', 'Q17')))

def calc_Segment_Revenue_Model_Q41(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Q12')), xl_ref(ctx.cell('Segment Revenue Model', 'Q17')))

def calc_Segment_Revenue_Model_Q42(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Q13')), xl_ref(ctx.cell('Segment Revenue Model', 'Q17')))

def calc_Segment_Revenue_Model_Q43(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Q14')), xl_ref(ctx.cell('Segment Revenue Model', 'Q17')))

def calc_Segment_Revenue_Model_Q44(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Q15')), xl_ref(ctx.cell('Segment Revenue Model', 'Q17')))

def calc_Segment_Revenue_Model_Q45(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Q16')), xl_ref(ctx.cell('Segment Revenue Model', 'Q17')))

def calc_Segment_Revenue_Model_Q46(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Q17')), xl_ref(ctx.cell('Segment Revenue Model', 'Q17')))

def calc_Segment_Revenue_Model_R33(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'R17')), xl_ref(ctx.cell('Segment Revenue Model', 'M17'))), 1)

def calc_Segment_Revenue_Model_R37(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'R8')), xl_ref(ctx.cell('Segment Revenue Model', 'R17')))

def calc_Segment_Revenue_Model_R38(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'R9')), xl_ref(ctx.cell('Segment Revenue Model', 'R17')))

def calc_Segment_Revenue_Model_R39(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'R10')), xl_ref(ctx.cell('Segment Revenue Model', 'R17')))

def calc_Segment_Revenue_Model_R40(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'R11')), xl_ref(ctx.cell('Segment Revenue Model', 'R17')))

def calc_Segment_Revenue_Model_R41(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'R12')), xl_ref(ctx.cell('Segment Revenue Model', 'R17')))

def calc_Segment_Revenue_Model_R42(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'R13')), xl_ref(ctx.cell('Segment Revenue Model', 'R17')))

def calc_Segment_Revenue_Model_R43(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'R14')), xl_ref(ctx.cell('Segment Revenue Model', 'R17')))

def calc_Segment_Revenue_Model_R44(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'R15')), xl_ref(ctx.cell('Segment Revenue Model', 'R17')))

def calc_Segment_Revenue_Model_R45(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'R16')), xl_ref(ctx.cell('Segment Revenue Model', 'R17')))

def calc_Segment_Revenue_Model_R46(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'R17')), xl_ref(ctx.cell('Segment Revenue Model', 'R17')))

def calc_Segment_Revenue_Model_S33(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'S17')), xl_ref(ctx.cell('Segment Revenue Model', 'N17'))), 1)

def calc_Segment_Revenue_Model_S37(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'S8')), xl_ref(ctx.cell('Segment Revenue Model', 'S17')))

def calc_Segment_Revenue_Model_S38(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'S9')), xl_ref(ctx.cell('Segment Revenue Model', 'S17')))

def calc_Segment_Revenue_Model_S39(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'S10')), xl_ref(ctx.cell('Segment Revenue Model', 'S17')))

def calc_Segment_Revenue_Model_S40(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'S11')), xl_ref(ctx.cell('Segment Revenue Model', 'S17')))

def calc_Segment_Revenue_Model_S41(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'S12')), xl_ref(ctx.cell('Segment Revenue Model', 'S17')))

def calc_Segment_Revenue_Model_S42(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'S13')), xl_ref(ctx.cell('Segment Revenue Model', 'S17')))

def calc_Segment_Revenue_Model_S43(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'S14')), xl_ref(ctx.cell('Segment Revenue Model', 'S17')))

def calc_Segment_Revenue_Model_S44(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'S15')), xl_ref(ctx.cell('Segment Revenue Model', 'S17')))

def calc_Segment_Revenue_Model_S45(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'S16')), xl_ref(ctx.cell('Segment Revenue Model', 'S17')))

def calc_Segment_Revenue_Model_S46(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'S17')), xl_ref(ctx.cell('Segment Revenue Model', 'S17')))

def calc_Segment_Revenue_Model_W16(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'W32')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'V16')))

def calc_Segment_Revenue_Model_F61(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'F58')), xl_ref(ctx.cell('Segment Revenue Model', 'F59'))), xl_ref(ctx.cell('Segment Revenue Model', 'F60')))

def calc_Segment_Revenue_Model_P61(ctx):
    return xl_add(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'P58')), xl_ref(ctx.cell('Segment Revenue Model', 'P59'))), xl_ref(ctx.cell('Segment Revenue Model', 'P60')))

def calc_Segment_Revenue_Model_U61(ctx):
    return xl_add(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'U58')), xl_ref(ctx.cell('Segment Revenue Model', 'U59'))), xl_ref(ctx.cell('Segment Revenue Model', 'U60')))

def calc_Segment_Revenue_Model_K61(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'G61:J61'))

def calc_Segment_Revenue_Model_K74(ctx):
    return xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'K72')), xl_ref(ctx.cell('Segment Revenue Model', 'K73')))

def calc_Segment_Revenue_Model_P74(ctx):
    return xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'P72')), xl_ref(ctx.cell('Segment Revenue Model', 'P73')))

def calc_Segment_Revenue_Model_F74(ctx):
    return xl_add(xl_add(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'B74')), xl_ref(ctx.cell('Segment Revenue Model', 'C74'))), xl_ref(ctx.cell('Segment Revenue Model', 'D74'))), xl_ref(ctx.cell('Segment Revenue Model', 'E74')))

def calc_Segment_Revenue_Model_U74(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'U72:U73'))

def calc_INCOME_STATEMENT_F7(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'F6')), xl_ref(ctx.cell('INCOME STATEMENT', 'H6'))), 100)

def calc_INCOME_STATEMENT_O6(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'M6')), xl_ref(ctx.cell('INCOME STATEMENT', 'N6')))

def calc_INCOME_STATEMENT_M8(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'M6')), xl_ref(ctx.cell('INCOME STATEMENT', 'F6'))), 1), 100)

def calc_INCOME_STATEMENT_T7(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'T6')), xl_ref(ctx.cell('INCOME STATEMENT', 'V6'))), 100)

def calc_INCOME_STATEMENT_T8(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'T6')), xl_ref(ctx.cell('INCOME STATEMENT', 'M6'))), 1), 100)

def calc_INCOME_STATEMENT_AC6(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AA6')), xl_ref(ctx.cell('INCOME STATEMENT', 'AB6')))

def calc_INCOME_STATEMENT_AA8(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AA6')), xl_ref(ctx.cell('INCOME STATEMENT', 'T6'))), 1), 100)

def calc_INCOME_STATEMENT_B25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'B21')), xl_ref(ctx.cell('INCOME STATEMENT', 'B6'))), 100)

def calc_INCOME_STATEMENT_B28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'B21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'B26')), xl_ref(ctx.cell('INCOME STATEMENT', 'B27'))))

def calc_INCOME_STATEMENT_C24(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'C21')), xl_ref(ctx.cell('INCOME STATEMENT', 'B21'))), 1), 100)

def calc_INCOME_STATEMENT_C25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'C21')), xl_ref(ctx.cell('INCOME STATEMENT', 'C6'))), 100)

def calc_INCOME_STATEMENT_C28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'C21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'C26')), xl_ref(ctx.cell('INCOME STATEMENT', 'C27'))))

def calc_INCOME_STATEMENT_E24(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'E21')), xl_ref(ctx.cell('INCOME STATEMENT', 'C21'))), 1), 100)

def calc_INCOME_STATEMENT_E25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'E21')), xl_ref(ctx.cell('INCOME STATEMENT', 'E6'))), 100)

def calc_INCOME_STATEMENT_E28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'E21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'E26')), xl_ref(ctx.cell('INCOME STATEMENT', 'E27'))))

def calc_INCOME_STATEMENT_G24(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'G21')), xl_ref(ctx.cell('INCOME STATEMENT', 'E21'))), 1), 100)

def calc_INCOME_STATEMENT_G25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'G21')), xl_ref(ctx.cell('INCOME STATEMENT', 'G6'))), 100)

def calc_INCOME_STATEMENT_G28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'G21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'G26')), xl_ref(ctx.cell('INCOME STATEMENT', 'G27'))))

def calc_INCOME_STATEMENT_I23(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'I21')), xl_ref(ctx.cell('INCOME STATEMENT', 'B21'))), 1), 100)

def calc_INCOME_STATEMENT_I24(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'I21')), xl_ref(ctx.cell('INCOME STATEMENT', 'G21'))), 1), 100)

def calc_INCOME_STATEMENT_I25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'I21')), xl_ref(ctx.cell('INCOME STATEMENT', 'I6'))), 100)

def calc_INCOME_STATEMENT_I28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'I21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'I26')), xl_ref(ctx.cell('INCOME STATEMENT', 'I27'))))

def calc_INCOME_STATEMENT_J23(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'J21')), xl_ref(ctx.cell('INCOME STATEMENT', 'C21'))), 1), 100)

def calc_INCOME_STATEMENT_J24(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'J21')), xl_ref(ctx.cell('INCOME STATEMENT', 'I21'))), 1), 100)

def calc_INCOME_STATEMENT_J25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'J21')), xl_ref(ctx.cell('INCOME STATEMENT', 'J6'))), 100)

def calc_INCOME_STATEMENT_J28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'J21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'J26')), xl_ref(ctx.cell('INCOME STATEMENT', 'J27'))))

def calc_INCOME_STATEMENT_L23(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'L21')), xl_ref(ctx.cell('INCOME STATEMENT', 'E21'))), 1), 100)

def calc_INCOME_STATEMENT_L24(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'L21')), xl_ref(ctx.cell('INCOME STATEMENT', 'J21'))), 1), 100)

def calc_INCOME_STATEMENT_L25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'L21')), xl_ref(ctx.cell('INCOME STATEMENT', 'L6'))), 100)

def calc_INCOME_STATEMENT_L28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'L21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'L26')), xl_ref(ctx.cell('INCOME STATEMENT', 'L27'))))

def calc_INCOME_STATEMENT_N23(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'N21')), xl_ref(ctx.cell('INCOME STATEMENT', 'G21'))), 1), 100)

def calc_INCOME_STATEMENT_N24(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'N21')), xl_ref(ctx.cell('INCOME STATEMENT', 'L21'))), 1), 100)

def calc_INCOME_STATEMENT_N25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'N21')), xl_ref(ctx.cell('INCOME STATEMENT', 'N6'))), 100)

def calc_INCOME_STATEMENT_N28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'N21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'N26')), xl_ref(ctx.cell('INCOME STATEMENT', 'N27'))))

def calc_INCOME_STATEMENT_P23(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'P21')), xl_ref(ctx.cell('INCOME STATEMENT', 'I21'))), 1), 100)

def calc_INCOME_STATEMENT_P24(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'P21')), xl_ref(ctx.cell('INCOME STATEMENT', 'N21'))), 1), 100)

def calc_INCOME_STATEMENT_P25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'P21')), xl_ref(ctx.cell('INCOME STATEMENT', 'P6'))), 100)

def calc_INCOME_STATEMENT_P28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'P21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'P26')), xl_ref(ctx.cell('INCOME STATEMENT', 'P27'))))

def calc_INCOME_STATEMENT_Q23(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Q21')), xl_ref(ctx.cell('INCOME STATEMENT', 'J21'))), 1), 100)

def calc_INCOME_STATEMENT_Q24(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Q21')), xl_ref(ctx.cell('INCOME STATEMENT', 'P21'))), 1), 100)

def calc_INCOME_STATEMENT_Q25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Q21')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q6'))), 100)

def calc_INCOME_STATEMENT_Q28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'Q21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'Q26')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q27'))))

def calc_INCOME_STATEMENT_S23(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'S21')), xl_ref(ctx.cell('INCOME STATEMENT', 'L21'))), 1), 100)

def calc_INCOME_STATEMENT_S24(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'S21')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q21'))), 1), 100)

def calc_INCOME_STATEMENT_S25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'S21')), xl_ref(ctx.cell('INCOME STATEMENT', 'S6'))), 100)

def calc_INCOME_STATEMENT_S28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'S21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'S26')), xl_ref(ctx.cell('INCOME STATEMENT', 'S27'))))

def calc_INCOME_STATEMENT_U23(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'U21')), xl_ref(ctx.cell('INCOME STATEMENT', 'N21'))), 1), 100)

def calc_INCOME_STATEMENT_U24(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'U21')), xl_ref(ctx.cell('INCOME STATEMENT', 'S21'))), 1), 100)

def calc_INCOME_STATEMENT_U25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'U21')), xl_ref(ctx.cell('INCOME STATEMENT', 'U6'))), 100)

def calc_INCOME_STATEMENT_U28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'U21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'U26')), xl_ref(ctx.cell('INCOME STATEMENT', 'U27'))))

def calc_INCOME_STATEMENT_W23(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'W21')), xl_ref(ctx.cell('INCOME STATEMENT', 'P21'))), 1), 100)

def calc_INCOME_STATEMENT_W24(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'W21')), xl_ref(ctx.cell('INCOME STATEMENT', 'U21'))), 1), 100)

def calc_INCOME_STATEMENT_W25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'W21')), xl_ref(ctx.cell('INCOME STATEMENT', 'W6'))), 100)

def calc_INCOME_STATEMENT_W28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'W21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'W26')), xl_ref(ctx.cell('INCOME STATEMENT', 'W27'))))

def calc_INCOME_STATEMENT_X23(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'X21')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q21'))), 1), 100)

def calc_INCOME_STATEMENT_X24(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'X21')), xl_ref(ctx.cell('INCOME STATEMENT', 'W21'))), 1), 100)

def calc_INCOME_STATEMENT_X25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'X21')), xl_ref(ctx.cell('INCOME STATEMENT', 'X6'))), 100)

def calc_INCOME_STATEMENT_X28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'X21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'X26')), xl_ref(ctx.cell('INCOME STATEMENT', 'X27'))))

def calc_INCOME_STATEMENT_Z23(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Z21')), xl_ref(ctx.cell('INCOME STATEMENT', 'S21'))), 1), 100)

def calc_INCOME_STATEMENT_Z24(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Z21')), xl_ref(ctx.cell('INCOME STATEMENT', 'X21'))), 1), 100)

def calc_INCOME_STATEMENT_Z25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Z21')), xl_ref(ctx.cell('INCOME STATEMENT', 'Z6'))), 100)

def calc_INCOME_STATEMENT_Z28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'Z21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'Z26')), xl_ref(ctx.cell('INCOME STATEMENT', 'Z27'))))

def calc_INCOME_STATEMENT_AB23(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AB21')), xl_ref(ctx.cell('INCOME STATEMENT', 'U21'))), 1), 100)

def calc_INCOME_STATEMENT_AB24(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AB21')), xl_ref(ctx.cell('INCOME STATEMENT', 'Z21'))), 1), 100)

def calc_INCOME_STATEMENT_AB25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AB21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AB6'))), 100)

def calc_INCOME_STATEMENT_AB28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AB21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AB26')), xl_ref(ctx.cell('INCOME STATEMENT', 'AB27'))))

def calc_INCOME_STATEMENT_F12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'F11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'F6'))))

def calc_INCOME_STATEMENT_O11(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'M11')), xl_ref(ctx.cell('INCOME STATEMENT', 'N11')))

def calc_INCOME_STATEMENT_M12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'M11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'M6'))))

def calc_INCOME_STATEMENT_T12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'T11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'T6'))))

def calc_INCOME_STATEMENT_AC11(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AA11')), xl_ref(ctx.cell('INCOME STATEMENT', 'AB11')))

def calc_INCOME_STATEMENT_AA12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AA11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AA6'))))

def calc_INCOME_STATEMENT_F14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'F13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'F6'))))

def calc_INCOME_STATEMENT_M14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'M13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'M6'))))

def calc_INCOME_STATEMENT_T14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'T13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'T6'))))

def calc_INCOME_STATEMENT_AC13(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AA13')), xl_ref(ctx.cell('INCOME STATEMENT', 'AB13')))

def calc_INCOME_STATEMENT_AA14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AA13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AA6'))))

def calc_INCOME_STATEMENT_H15(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'F15')), xl_ref(ctx.cell('INCOME STATEMENT', 'G15')))

def calc_INCOME_STATEMENT_F16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'F15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'F6'))))

def calc_INCOME_STATEMENT_O15(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'M15')), xl_ref(ctx.cell('INCOME STATEMENT', 'N15')))

def calc_INCOME_STATEMENT_M16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'M15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'M6'))))

def calc_INCOME_STATEMENT_V15(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'T15')), xl_ref(ctx.cell('INCOME STATEMENT', 'U15')))

def calc_INCOME_STATEMENT_T16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'T15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'T6'))))

def calc_INCOME_STATEMENT_AC15(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AA15')), xl_ref(ctx.cell('INCOME STATEMENT', 'AB15')))

def calc_INCOME_STATEMENT_AA16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AA15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AA6'))))

def calc_INCOME_STATEMENT_F18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'F17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'F6'))))

def calc_INCOME_STATEMENT_O17(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'M17')), xl_ref(ctx.cell('INCOME STATEMENT', 'N17')))

def calc_INCOME_STATEMENT_M18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'M17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'M6'))))

def calc_INCOME_STATEMENT_T18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'T17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'T6'))))

def calc_INCOME_STATEMENT_AC17(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AA17')), xl_ref(ctx.cell('INCOME STATEMENT', 'AB17')))

def calc_INCOME_STATEMENT_AA18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AA17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AA6'))))

def calc_INCOME_STATEMENT_D21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'D6')), xl_ref(ctx.cell('INCOME STATEMENT', 'D10')))

def calc_INCOME_STATEMENT_F10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'F11')), xl_ref(ctx.cell('INCOME STATEMENT', 'F13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'F15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'F17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'F19')))

def calc_INCOME_STATEMENT_H19(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'F19')), xl_ref(ctx.cell('INCOME STATEMENT', 'G19')))

def calc_INCOME_STATEMENT_F20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'F19')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'F6'))))

def calc_INCOME_STATEMENT_M10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'M11')), xl_ref(ctx.cell('INCOME STATEMENT', 'M13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'M17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'M15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'M19')))

def calc_INCOME_STATEMENT_M20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'M19')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'M6'))))

def calc_INCOME_STATEMENT_R21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'R6')), xl_ref(ctx.cell('INCOME STATEMENT', 'R10')))

def calc_INCOME_STATEMENT_T10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'T11')), xl_ref(ctx.cell('INCOME STATEMENT', 'T13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'T17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'T15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'T19')))

def calc_INCOME_STATEMENT_V19(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'T19')), xl_ref(ctx.cell('INCOME STATEMENT', 'U19')))

def calc_INCOME_STATEMENT_T20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'T19')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'T6'))))

def calc_INCOME_STATEMENT_Y21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'Y6')), xl_ref(ctx.cell('INCOME STATEMENT', 'Y10')))

def calc_INCOME_STATEMENT_AA10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AA11')), xl_ref(ctx.cell('INCOME STATEMENT', 'AA13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AA17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AA15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AA19')))

def calc_INCOME_STATEMENT_AC19(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AA19')), xl_ref(ctx.cell('INCOME STATEMENT', 'AB19')))

def calc_INCOME_STATEMENT_AA20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AA19')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AA6'))))

def calc_PRESENTATION_G64(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AD6')), xl_ref(ctx.cell('INCOME STATEMENT', 'AD10')))

def calc_INCOME_STATEMENT_AD21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AD6')), xl_ref(ctx.cell('INCOME STATEMENT', 'AD10')))

def calc_PRESENTATION_H64(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AE6')), xl_ref(ctx.cell('INCOME STATEMENT', 'AE10')))

def calc_INCOME_STATEMENT_AE21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AE6')), xl_ref(ctx.cell('INCOME STATEMENT', 'AE10')))

def calc_PRESENTATION_I64(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AF6')), xl_ref(ctx.cell('INCOME STATEMENT', 'AF10')))

def calc_INCOME_STATEMENT_AF21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AF6')), xl_ref(ctx.cell('INCOME STATEMENT', 'AF10')))

def calc_PRESENTATION_J64(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AG6')), xl_ref(ctx.cell('INCOME STATEMENT', 'AG10')))

def calc_INCOME_STATEMENT_AG21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AG6')), xl_ref(ctx.cell('INCOME STATEMENT', 'AG10')))

def calc_INCOME_STATEMENT_H26(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'F26')), xl_ref(ctx.cell('INCOME STATEMENT', 'G26')))

def calc_INCOME_STATEMENT_O26(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'M26')), xl_ref(ctx.cell('INCOME STATEMENT', 'N26')))

def calc_INCOME_STATEMENT_V26(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'T26')), xl_ref(ctx.cell('INCOME STATEMENT', 'U26')))

def calc_INCOME_STATEMENT_AC26(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AA26')), xl_ref(ctx.cell('INCOME STATEMENT', 'AB26')))

def calc_INCOME_STATEMENT_H27(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'F27')), xl_ref(ctx.cell('INCOME STATEMENT', 'G27')))

def calc_INCOME_STATEMENT_O27(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'M27')), xl_ref(ctx.cell('INCOME STATEMENT', 'N27')))

def calc_INCOME_STATEMENT_K30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'K28')), xl_ref(ctx.cell('INCOME STATEMENT', 'K29')))

def calc_INCOME_STATEMENT_V27(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'T27')), xl_ref(ctx.cell('INCOME STATEMENT', 'U27')))

def calc_INCOME_STATEMENT_H29(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'F29')), xl_ref(ctx.cell('INCOME STATEMENT', 'G29')))

def calc_INCOME_STATEMENT_O29(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'M29')), xl_ref(ctx.cell('INCOME STATEMENT', 'N29')))

def calc_INCOME_STATEMENT_AC29(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AA29')), xl_ref(ctx.cell('INCOME STATEMENT', 'AB29')))

def calc_INCOME_STATEMENT_AC33(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AA33')), xl_ref(ctx.cell('INCOME STATEMENT', 'AB33')))

def calc_INCOME_STATEMENT_H34(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'F34')), xl_ref(ctx.cell('INCOME STATEMENT', 'G34')))

def calc_INCOME_STATEMENT_AC34(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AA34')), xl_ref(ctx.cell('INCOME STATEMENT', 'AB34')))

def calc_INCOME_STATEMENT_H35(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'F35')), xl_ref(ctx.cell('INCOME STATEMENT', 'G35')))

def calc_INCOME_STATEMENT_E47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'D47'))

def calc_INCOME_STATEMENT_D48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'D47')), 10)

def calc_INCOME_STATEMENT_V47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'U47'))

def calc_INCOME_STATEMENT_U48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'U47')), 10)

def calc_CASH_FOW_STATEMENT_I12(ctx):
    return xl_add(xl_ref(ctx.cell('CASH FOW STATEMENT', 'I10')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'I11')))

def calc_Valuation_B46(ctx):
    return xl_div(xl_uminus(xl_ref(ctx.cell('Valuation', 'B27'))), xl_ref(ctx.cell('Valuation', 'B19')))

def calc_Valuation_D46(ctx):
    return xl_div(xl_uminus(xl_ref(ctx.cell('Valuation', 'D27'))), xl_ref(ctx.cell('Valuation', 'D19')))

def calc_Valuation_E46(ctx):
    return xl_div(xl_uminus(xl_ref(ctx.cell('Valuation', 'E27'))), xl_ref(ctx.cell('Valuation', 'E19')))

def calc_Valuation_H25(ctx):
    return xl_ref(ctx.cell('Valuation', 'G25'))

def calc_Ratio_Analysis_B8(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'B19')), xl_ref(ctx.cell('BALANCESHEET', 'B8')))

def calc_Ratio_Analysis_C8(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'C19')), xl_ref(ctx.cell('BALANCESHEET', 'C8')))

def calc_Ratio_Analysis_D8(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'D19')), xl_ref(ctx.cell('BALANCESHEET', 'D8')))

def calc_Ratio_Analysis_E8(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'E19')), xl_ref(ctx.cell('BALANCESHEET', 'E8')))

def calc_Ratio_Analysis_F8(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'F19')), xl_ref(ctx.cell('BALANCESHEET', 'F8')))

def calc_Ratio_Analysis_G8(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'G19')), xl_ref(ctx.cell('BALANCESHEET', 'G8')))

def calc_Ratio_Analysis_H8(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'H19')), xl_ref(ctx.cell('BALANCESHEET', 'H8')))

def calc_Ratio_Analysis_I8(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'I19')), xl_ref(ctx.cell('BALANCESHEET', 'I8')))

def calc_Ratio_Analysis_J8(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'J19')), xl_ref(ctx.cell('BALANCESHEET', 'J8')))

def calc_Ratio_Analysis_B7(ctx):
    return xl_div(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'B19')), xl_ref(ctx.cell('BALANCESHEET', 'B23'))), xl_percent(xl_ref(ctx.cell('BALANCESHEET', 'B24'))))

def calc_Ratio_Analysis_C7(ctx):
    return xl_div(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'C19')), xl_ref(ctx.cell('BALANCESHEET', 'C23'))), xl_percent(xl_ref(ctx.cell('BALANCESHEET', 'C24'))))

def calc_Ratio_Analysis_D7(ctx):
    return xl_div(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'D19')), xl_ref(ctx.cell('BALANCESHEET', 'D23'))), xl_percent(xl_ref(ctx.cell('BALANCESHEET', 'D24'))))

def calc_Ratio_Analysis_E7(ctx):
    return xl_div(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'E19')), xl_ref(ctx.cell('BALANCESHEET', 'E23'))), xl_percent(xl_ref(ctx.cell('BALANCESHEET', 'E24'))))

def calc_Valuation_G12(ctx):
    return xl_add(xl_ref(ctx.cell('Valuation', 'G10')), xl_ref(ctx.cell('Valuation', 'G11')))

def calc_Valuation_G53(ctx):
    return xl_uminus(xl_ref(ctx.cell('Valuation', 'G10')))

def calc_Ratio_Analysis_F7(ctx):
    return xl_div(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'F19')), xl_ref(ctx.cell('BALANCESHEET', 'F23'))), xl_percent(xl_ref(ctx.cell('BALANCESHEET', 'F24'))))

def calc_Ratio_Analysis_G7(ctx):
    return xl_div(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'G19')), xl_ref(ctx.cell('BALANCESHEET', 'G23'))), xl_percent(xl_ref(ctx.cell('BALANCESHEET', 'G24'))))

def calc_Ratio_Analysis_H7(ctx):
    return xl_div(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'H19')), xl_ref(ctx.cell('BALANCESHEET', 'H23'))), xl_percent(xl_ref(ctx.cell('BALANCESHEET', 'H24'))))

def calc_Ratio_Analysis_I7(ctx):
    return xl_div(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'I19')), xl_ref(ctx.cell('BALANCESHEET', 'I23'))), xl_percent(xl_ref(ctx.cell('BALANCESHEET', 'I24'))))

def calc_Ratio_Analysis_J7(ctx):
    return xl_div(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'J19')), xl_ref(ctx.cell('BALANCESHEET', 'J23'))), xl_percent(xl_ref(ctx.cell('BALANCESHEET', 'J24'))))

def calc_Ratio_Analysis_B17(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'B39')), xl_ref(ctx.cell('BALANCESHEET', 'B27')))

def calc_Ratio_Analysis_C10(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'H6')), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'C27')), xl_ref(ctx.cell('BALANCESHEET', 'B27'))))

def calc_Ratio_Analysis_C17(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'C39')), xl_ref(ctx.cell('BALANCESHEET', 'C27')))

def calc_Ratio_Analysis_D17(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'D39')), xl_ref(ctx.cell('BALANCESHEET', 'D27')))

def calc_Ratio_Analysis_E10(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'V6')), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'E27')), xl_ref(ctx.cell('BALANCESHEET', 'D27'))))

def calc_Ratio_Analysis_E17(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'E39')), xl_ref(ctx.cell('BALANCESHEET', 'E27')))

def calc_BALANCESHEET_F60(ctx):
    return xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'F27')), xl_ref(ctx.cell('BALANCESHEET', 'F33'))), xl_ref(ctx.cell('BALANCESHEET', 'F34'))), xl_ref(ctx.cell('BALANCESHEET', 'F58')))

def calc_Ratio_Analysis_F17(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'F39')), xl_ref(ctx.cell('BALANCESHEET', 'F27')))

def calc_BALANCESHEET_G60(ctx):
    return xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'G27')), xl_ref(ctx.cell('BALANCESHEET', 'G33'))), xl_ref(ctx.cell('BALANCESHEET', 'G34'))), xl_ref(ctx.cell('BALANCESHEET', 'G58')))

def calc_Ratio_Analysis_G10(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AD6')), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'G27')), xl_ref(ctx.cell('BALANCESHEET', 'F27'))))

def calc_Ratio_Analysis_G17(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'G39')), xl_ref(ctx.cell('BALANCESHEET', 'G27')))

def calc_BALANCESHEET_H60(ctx):
    return xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'H27')), xl_ref(ctx.cell('BALANCESHEET', 'H33'))), xl_ref(ctx.cell('BALANCESHEET', 'H34'))), xl_ref(ctx.cell('BALANCESHEET', 'H58')))

def calc_Ratio_Analysis_H10(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AE6')), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'H27')), xl_ref(ctx.cell('BALANCESHEET', 'G27'))))

def calc_Ratio_Analysis_H17(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'H39')), xl_ref(ctx.cell('BALANCESHEET', 'H27')))

def calc_BALANCESHEET_I60(ctx):
    return xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'I27')), xl_ref(ctx.cell('BALANCESHEET', 'I33'))), xl_ref(ctx.cell('BALANCESHEET', 'I34'))), xl_ref(ctx.cell('BALANCESHEET', 'I58')))

def calc_Ratio_Analysis_I10(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AF6')), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'I27')), xl_ref(ctx.cell('BALANCESHEET', 'H27'))))

def calc_Ratio_Analysis_I17(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'I39')), xl_ref(ctx.cell('BALANCESHEET', 'I27')))

def calc_BALANCESHEET_J60(ctx):
    return xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'J27')), xl_ref(ctx.cell('BALANCESHEET', 'J33'))), xl_ref(ctx.cell('BALANCESHEET', 'J34'))), xl_ref(ctx.cell('BALANCESHEET', 'J58')))

def calc_Ratio_Analysis_J10(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AG6')), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'J27')), xl_ref(ctx.cell('BALANCESHEET', 'I27'))))

def calc_Ratio_Analysis_J17(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'J39')), xl_ref(ctx.cell('BALANCESHEET', 'J27')))

def calc_Ratio_Analysis_B9(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'B36')), xl_ref(ctx.cell('BALANCESHEET', 'B43')))

def calc_BALANCESHEET_B60(ctx):
    return xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'B27')), xl_ref(ctx.cell('BALANCESHEET', 'B33'))), xl_ref(ctx.cell('BALANCESHEET', 'B34'))), xl_ref(ctx.cell('BALANCESHEET', 'B58')))

def calc_CASH_FOW_STATEMENT_B20(ctx):
    return xl_sub(xl_sub(xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'C37')), xl_ref(ctx.cell('BALANCESHEET', 'C38'))), xl_ref(ctx.cell('BALANCESHEET', 'C40'))), xl_ref(ctx.cell('BALANCESHEET', 'C41'))), xl_ref(ctx.cell('BALANCESHEET', 'C43'))), xl_sub(xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'B37')), xl_ref(ctx.cell('BALANCESHEET', 'B38'))), xl_ref(ctx.cell('BALANCESHEET', 'B40'))), xl_ref(ctx.cell('BALANCESHEET', 'B41'))), xl_ref(ctx.cell('BALANCESHEET', 'B43'))))

def calc_Ratio_Analysis_C9(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'C36')), xl_ref(ctx.cell('BALANCESHEET', 'C43')))

def calc_BALANCESHEET_C60(ctx):
    return xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'C27')), xl_ref(ctx.cell('BALANCESHEET', 'C33'))), xl_ref(ctx.cell('BALANCESHEET', 'C34'))), xl_ref(ctx.cell('BALANCESHEET', 'C58')))

def calc_Ratio_Analysis_C24(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'C58')), xl_ref(ctx.cell('INCOME STATEMENT', 'H6')))

def calc_CASH_FOW_STATEMENT_C20(ctx):
    return xl_sub(xl_sub(xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'D37')), xl_ref(ctx.cell('BALANCESHEET', 'D38'))), xl_ref(ctx.cell('BALANCESHEET', 'D40'))), xl_ref(ctx.cell('BALANCESHEET', 'D41'))), xl_ref(ctx.cell('BALANCESHEET', 'D43'))), xl_sub(xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'C37')), xl_ref(ctx.cell('BALANCESHEET', 'C38'))), xl_ref(ctx.cell('BALANCESHEET', 'C40'))), xl_ref(ctx.cell('BALANCESHEET', 'C41'))), xl_ref(ctx.cell('BALANCESHEET', 'C43'))))

def calc_Ratio_Analysis_D9(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'D36')), xl_ref(ctx.cell('BALANCESHEET', 'D43')))

def calc_BALANCESHEET_D60(ctx):
    return xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'D27')), xl_ref(ctx.cell('BALANCESHEET', 'D33'))), xl_ref(ctx.cell('BALANCESHEET', 'D34'))), xl_ref(ctx.cell('BALANCESHEET', 'D58')))

def calc_CASH_FOW_STATEMENT_D20(ctx):
    return xl_sub(xl_sub(xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'E37')), xl_ref(ctx.cell('BALANCESHEET', 'E38'))), xl_ref(ctx.cell('BALANCESHEET', 'E40'))), xl_ref(ctx.cell('BALANCESHEET', 'E41'))), xl_ref(ctx.cell('BALANCESHEET', 'E43'))), xl_sub(xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'D37')), xl_ref(ctx.cell('BALANCESHEET', 'D38'))), xl_ref(ctx.cell('BALANCESHEET', 'D40'))), xl_ref(ctx.cell('BALANCESHEET', 'D41'))), xl_ref(ctx.cell('BALANCESHEET', 'D43'))))

def calc_CASH_FOW_STATEMENT_E20(ctx):
    return xl_sub(xl_sub(xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'F37')), xl_ref(ctx.cell('BALANCESHEET', 'F38'))), xl_ref(ctx.cell('BALANCESHEET', 'F40'))), xl_ref(ctx.cell('BALANCESHEET', 'F41'))), xl_ref(ctx.cell('BALANCESHEET', 'F43'))), xl_sub(xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'E37')), xl_ref(ctx.cell('BALANCESHEET', 'E38'))), xl_ref(ctx.cell('BALANCESHEET', 'E40'))), xl_ref(ctx.cell('BALANCESHEET', 'E41'))), xl_ref(ctx.cell('BALANCESHEET', 'E43'))))

def calc_Ratio_Analysis_E9(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'E36')), xl_ref(ctx.cell('BALANCESHEET', 'E43')))

def calc_BALANCESHEET_E60(ctx):
    return xl_add(xl_add(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'E27')), xl_ref(ctx.cell('BALANCESHEET', 'E33'))), xl_ref(ctx.cell('BALANCESHEET', 'E34'))), xl_ref(ctx.cell('BALANCESHEET', 'E58')))

def calc_Ratio_Analysis_E24(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'E58')), xl_ref(ctx.cell('INCOME STATEMENT', 'V6')))

def calc_Valuation_F37(ctx):
    return xl_sub(xl_ref(ctx.cell('Valuation', 'E37')), xl_div(xl_sub(xl_ref(ctx.cell('Valuation', 'E37')), xl_ref(ctx.cell('Valuation', 'J37'))), 5))

def calc_Valuation_F45(ctx):
    return xl_ref(ctx.cell('Valuation', 'E45'))

def calc_PRESENTATION_U17(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'U15')), xl_ref(ctx.cell('PRESENTATION', 'U16')))

def calc_PRESENTATION_B38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'B36')), xl_ref(ctx.cell('PRESENTATION', 'B37')))

def calc_PRESENTATION_C38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'C36')), xl_ref(ctx.cell('PRESENTATION', 'C37')))

def calc_PRESENTATION_E38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'E36')), xl_ref(ctx.cell('PRESENTATION', 'E37')))

def calc_PRESENTATION_G38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'G36')), xl_ref(ctx.cell('PRESENTATION', 'G37')))

def calc_PRESENTATION_I38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'I36')), xl_ref(ctx.cell('PRESENTATION', 'I37')))

def calc_PRESENTATION_J38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'J36')), xl_ref(ctx.cell('PRESENTATION', 'J37')))

def calc_PRESENTATION_L38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'L36')), xl_ref(ctx.cell('PRESENTATION', 'L37')))

def calc_PRESENTATION_N38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'N36')), xl_ref(ctx.cell('PRESENTATION', 'N37')))

def calc_PRESENTATION_P38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'P36')), xl_ref(ctx.cell('PRESENTATION', 'P37')))

def calc_PRESENTATION_Q38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'Q36')), xl_ref(ctx.cell('PRESENTATION', 'Q37')))

def calc_PRESENTATION_S38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'S36')), xl_ref(ctx.cell('PRESENTATION', 'S37')))

def calc_PRESENTATION_U38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'U36')), xl_ref(ctx.cell('PRESENTATION', 'U37')))

def calc_PRESENTATION_W38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'W36')), xl_ref(ctx.cell('PRESENTATION', 'W37')))

def calc_PRESENTATION_X38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'X36')), xl_ref(ctx.cell('PRESENTATION', 'X37')))

def calc_PRESENTATION_Z38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'Z36')), xl_ref(ctx.cell('PRESENTATION', 'Z37')))

def calc_PRESENTATION_AB38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'AB36')), xl_ref(ctx.cell('PRESENTATION', 'AB37')))

def calc_PRESENTATION_AD38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'AD36')), xl_ref(ctx.cell('PRESENTATION', 'AD37')))

def calc_PRESENTATION_AE38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'AE36')), xl_ref(ctx.cell('PRESENTATION', 'AE37')))

def calc_PRESENTATION_AF38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'AF36')), xl_ref(ctx.cell('PRESENTATION', 'AF37')))

def calc_PRESENTATION_AG38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'AG36')), xl_ref(ctx.cell('PRESENTATION', 'AG37')))

def calc_PRESENTATION_H26(ctx):
    return xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'H27')), xl_ref(ctx.cell('PRESENTATION', 'H28'))), xl_ref(ctx.cell('PRESENTATION', 'H30'))), xl_ref(ctx.cell('PRESENTATION', 'H29')))

def calc_PRESENTATION_O26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'O27')), xl_ref(ctx.cell('PRESENTATION', 'O28'))), xl_ref(ctx.cell('PRESENTATION', 'O30'))), xl_ref(ctx.cell('PRESENTATION', 'O29'))), xl_ref(ctx.cell('PRESENTATION', 'O31')))

def calc_PRESENTATION_D33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'D32')), xl_ref(ctx.cell('PRESENTATION', 'D25'))), 100)

def calc_PRESENTATION_D36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'D32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'D34')), xl_ref(ctx.cell('PRESENTATION', 'D35'))))

def calc_PRESENTATION_F32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'F25')), xl_ref(ctx.cell('PRESENTATION', 'F26')))

def calc_PRESENTATION_M32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'M25')), xl_ref(ctx.cell('PRESENTATION', 'M26')))

def calc_PRESENTATION_R33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'R32')), xl_ref(ctx.cell('PRESENTATION', 'R25'))), 100)

def calc_PRESENTATION_R36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'R32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'R34')), xl_ref(ctx.cell('PRESENTATION', 'R35'))))

def calc_PRESENTATION_T32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'T25')), xl_ref(ctx.cell('PRESENTATION', 'T26')))

def calc_PRESENTATION_V26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'V27')), xl_ref(ctx.cell('PRESENTATION', 'V28'))), xl_ref(ctx.cell('PRESENTATION', 'V30'))), xl_ref(ctx.cell('PRESENTATION', 'V29'))), xl_ref(ctx.cell('PRESENTATION', 'V31')))

def calc_PRESENTATION_Y33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'Y32')), xl_ref(ctx.cell('PRESENTATION', 'Y25'))), 100)

def calc_PRESENTATION_Y36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'Y32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'Y34')), xl_ref(ctx.cell('PRESENTATION', 'Y35'))))

def calc_PRESENTATION_AA32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AA25')), xl_ref(ctx.cell('PRESENTATION', 'AA26')))

def calc_PRESENTATION_AC26(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('PRESENTATION', 'AC27')), xl_ref(ctx.cell('PRESENTATION', 'AC28'))), xl_ref(ctx.cell('PRESENTATION', 'AC30'))), xl_ref(ctx.cell('PRESENTATION', 'AC29'))), xl_ref(ctx.cell('PRESENTATION', 'AC31')))

def calc_PRESENTATION_K40(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'K38')), xl_ref(ctx.cell('PRESENTATION', 'K39')))

def calc_Segment_Revenue_Model_T24(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'T8')), xl_ref(ctx.cell('Segment Revenue Model', 'O8'))), 1)

def calc_Segment_Revenue_Model_W8(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'W24')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'V8')))

def calc_Segment_Revenue_Model_T25(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'T9')), xl_ref(ctx.cell('Segment Revenue Model', 'O9'))), 1)

def calc_Segment_Revenue_Model_W9(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'W25')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'V9')))

def calc_Segment_Revenue_Model_T26(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'T10')), xl_ref(ctx.cell('Segment Revenue Model', 'O10'))), 1)

def calc_Segment_Revenue_Model_W10(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'W26')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'V10')))

def calc_Segment_Revenue_Model_T27(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'T11')), xl_ref(ctx.cell('Segment Revenue Model', 'O11'))), 1)

def calc_Segment_Revenue_Model_W11(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'W27')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'V11')))

def calc_Segment_Revenue_Model_X12(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'X28')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'W12')))

def calc_Segment_Revenue_Model_T29(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'T13')), xl_ref(ctx.cell('Segment Revenue Model', 'O13'))), 1)

def calc_Segment_Revenue_Model_W13(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'W29')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'V13')))

def calc_Segment_Revenue_Model_K37(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'K8')), xl_ref(ctx.cell('Segment Revenue Model', 'K17')))

def calc_Segment_Revenue_Model_K38(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'K9')), xl_ref(ctx.cell('Segment Revenue Model', 'K17')))

def calc_Segment_Revenue_Model_K39(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'K10')), xl_ref(ctx.cell('Segment Revenue Model', 'K17')))

def calc_Segment_Revenue_Model_K40(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'K11')), xl_ref(ctx.cell('Segment Revenue Model', 'K17')))

def calc_Segment_Revenue_Model_K41(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'K12')), xl_ref(ctx.cell('Segment Revenue Model', 'K17')))

def calc_Segment_Revenue_Model_K42(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'K13')), xl_ref(ctx.cell('Segment Revenue Model', 'K17')))

def calc_Segment_Revenue_Model_K43(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'K14')), xl_ref(ctx.cell('Segment Revenue Model', 'K17')))

def calc_Segment_Revenue_Model_K44(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'K15')), xl_ref(ctx.cell('Segment Revenue Model', 'K17')))

def calc_Segment_Revenue_Model_K45(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'K16')), xl_ref(ctx.cell('Segment Revenue Model', 'K17')))

def calc_Segment_Revenue_Model_K46(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'K17')), xl_ref(ctx.cell('Segment Revenue Model', 'K17')))

def calc_Segment_Revenue_Model_T15(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'T8:T14'))

def calc_Segment_Revenue_Model_T30(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'T14')), xl_ref(ctx.cell('Segment Revenue Model', 'O14'))), 1)

def calc_Segment_Revenue_Model_W14(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'W30')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'V14')))

def calc_Segment_Revenue_Model_V15(ctx):
    return xl_sum(ctx.range('Segment Revenue Model', 'V8:V14'))

def calc_Segment_Revenue_Model_P33(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'P17')), xl_ref(ctx.cell('Segment Revenue Model', 'K17'))), 1)

def calc_Segment_Revenue_Model_P37(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'P8')), xl_ref(ctx.cell('Segment Revenue Model', 'P17')))

def calc_Segment_Revenue_Model_P38(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'P9')), xl_ref(ctx.cell('Segment Revenue Model', 'P17')))

def calc_Segment_Revenue_Model_P39(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'P10')), xl_ref(ctx.cell('Segment Revenue Model', 'P17')))

def calc_Segment_Revenue_Model_P40(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'P11')), xl_ref(ctx.cell('Segment Revenue Model', 'P17')))

def calc_Segment_Revenue_Model_P41(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'P12')), xl_ref(ctx.cell('Segment Revenue Model', 'P17')))

def calc_Segment_Revenue_Model_P42(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'P13')), xl_ref(ctx.cell('Segment Revenue Model', 'P17')))

def calc_Segment_Revenue_Model_P43(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'P14')), xl_ref(ctx.cell('Segment Revenue Model', 'P17')))

def calc_Segment_Revenue_Model_P44(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'P15')), xl_ref(ctx.cell('Segment Revenue Model', 'P17')))

def calc_Segment_Revenue_Model_P45(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'P16')), xl_ref(ctx.cell('Segment Revenue Model', 'P17')))

def calc_Segment_Revenue_Model_P46(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'P17')), xl_ref(ctx.cell('Segment Revenue Model', 'P17')))

def calc_Segment_Revenue_Model_K33(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'K17')), xl_ref(ctx.cell('Segment Revenue Model', 'F17'))), 1)

def calc_Segment_Revenue_Model_X16(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'X32')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'W16')))

def calc_INCOME_STATEMENT_H7(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'F7')), xl_ref(ctx.cell('INCOME STATEMENT', 'G7')))

def calc_INCOME_STATEMENT_I7(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'I6')), xl_ref(ctx.cell('INCOME STATEMENT', 'O6'))), 100)

def calc_INCOME_STATEMENT_J7(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'J6')), xl_ref(ctx.cell('INCOME STATEMENT', 'O6'))), 100)

def calc_INCOME_STATEMENT_K7(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'K6')), xl_ref(ctx.cell('INCOME STATEMENT', 'O6'))), 100)

def calc_INCOME_STATEMENT_L7(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'L6')), xl_ref(ctx.cell('INCOME STATEMENT', 'O6'))), 100)

def calc_INCOME_STATEMENT_M7(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'M6')), xl_ref(ctx.cell('INCOME STATEMENT', 'O6'))), 100)

def calc_INCOME_STATEMENT_N7(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'N6')), xl_ref(ctx.cell('INCOME STATEMENT', 'O6'))), 100)

def calc_INCOME_STATEMENT_O8(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'O6')), xl_ref(ctx.cell('INCOME STATEMENT', 'H6'))), 1), 100)

def calc_INCOME_STATEMENT_V8(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'V6')), xl_ref(ctx.cell('INCOME STATEMENT', 'O6'))), 1), 100)

def calc_INCOME_STATEMENT_O14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'O13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'O6'))))

def calc_INCOME_STATEMENT_O20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'O19')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'O6'))))

def calc_Ratio_Analysis_D10(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'O6')), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'D27')), xl_ref(ctx.cell('BALANCESHEET', 'C27'))))

def calc_Ratio_Analysis_D21(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'O6')), xl_ref(ctx.cell('INCOME STATEMENT', 'H6'))), 1)

def calc_Ratio_Analysis_E21(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'V6')), xl_ref(ctx.cell('INCOME STATEMENT', 'O6'))), 1)

def calc_Ratio_Analysis_D23(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'O6')), xl_mul(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'D37')), xl_ref(ctx.cell('BALANCESHEET', 'C37'))), 0.5))

def calc_Ratio_Analysis_D24(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'D58')), xl_ref(ctx.cell('INCOME STATEMENT', 'O6')))

def calc_INCOME_STATEMENT_V7(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'T7')), xl_ref(ctx.cell('INCOME STATEMENT', 'U7')))

def calc_COMPANY_OVERVIEW_F28(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'G7')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC6')))

def calc_INCOME_STATEMENT_W7(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'W6')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC6'))), 100)

def calc_INCOME_STATEMENT_X7(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'X6')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC6'))), 100)

def calc_INCOME_STATEMENT_Y7(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Y6')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC6'))), 100)

def calc_INCOME_STATEMENT_Z7(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Z6')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC6'))), 100)

def calc_INCOME_STATEMENT_AA7(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AA6')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC6'))), 100)

def calc_INCOME_STATEMENT_AB7(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AB6')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC6'))), 100)

def calc_INCOME_STATEMENT_AC8(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AC6')), xl_ref(ctx.cell('INCOME STATEMENT', 'V6'))), 1), 100)

def calc_INCOME_STATEMENT_AD8(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AD6')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC6'))), 1), 100)

def calc_Valuation_C19(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AC6'))

def calc_Ratio_Analysis_F10(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AC6')), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'F27')), xl_ref(ctx.cell('BALANCESHEET', 'E27'))))

def calc_Ratio_Analysis_F21(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AC6')), xl_ref(ctx.cell('INCOME STATEMENT', 'V6'))), 1)

def calc_Ratio_Analysis_F23(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AC6')), xl_mul(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'F37')), xl_ref(ctx.cell('BALANCESHEET', 'E37'))), 0.5))

def calc_Ratio_Analysis_F24(ctx):
    return xl_div(xl_ref(ctx.cell('BALANCESHEET', 'F58')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC6')))

def calc_INCOME_STATEMENT_B30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'B28')), xl_ref(ctx.cell('INCOME STATEMENT', 'B29')))

def calc_INCOME_STATEMENT_C30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'C28')), xl_ref(ctx.cell('INCOME STATEMENT', 'C29')))

def calc_INCOME_STATEMENT_E30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'E28')), xl_ref(ctx.cell('INCOME STATEMENT', 'E29')))

def calc_INCOME_STATEMENT_G30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'G28')), xl_ref(ctx.cell('INCOME STATEMENT', 'G29')))

def calc_INCOME_STATEMENT_I30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'I28')), xl_ref(ctx.cell('INCOME STATEMENT', 'I29')))

def calc_INCOME_STATEMENT_J30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'J28')), xl_ref(ctx.cell('INCOME STATEMENT', 'J29')))

def calc_INCOME_STATEMENT_L30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'L28')), xl_ref(ctx.cell('INCOME STATEMENT', 'L29')))

def calc_INCOME_STATEMENT_N30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'N28')), xl_ref(ctx.cell('INCOME STATEMENT', 'N29')))

def calc_INCOME_STATEMENT_P30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'P28')), xl_ref(ctx.cell('INCOME STATEMENT', 'P29')))

def calc_INCOME_STATEMENT_Q30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'Q28')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q29')))

def calc_INCOME_STATEMENT_S30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'S28')), xl_ref(ctx.cell('INCOME STATEMENT', 'S29')))

def calc_INCOME_STATEMENT_U30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'U28')), xl_ref(ctx.cell('INCOME STATEMENT', 'U29')))

def calc_INCOME_STATEMENT_W30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'W28')), xl_ref(ctx.cell('INCOME STATEMENT', 'W29')))

def calc_INCOME_STATEMENT_X30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'X28')), xl_ref(ctx.cell('INCOME STATEMENT', 'X29')))

def calc_INCOME_STATEMENT_Z30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'Z28')), xl_ref(ctx.cell('INCOME STATEMENT', 'Z29')))

def calc_INCOME_STATEMENT_AB30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AB28')), xl_ref(ctx.cell('INCOME STATEMENT', 'AB29')))

def calc_INCOME_STATEMENT_O12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'O11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'O6'))))

def calc_INCOME_STATEMENT_AC12(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AC11')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AC6'))))

def calc_INCOME_STATEMENT_AC14(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AC13')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AC6'))))

def calc_INCOME_STATEMENT_H10(ctx):
    return xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'H11')), xl_ref(ctx.cell('INCOME STATEMENT', 'H13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'H17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'H15')))

def calc_INCOME_STATEMENT_H16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'H15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'H6'))))

def calc_INCOME_STATEMENT_O16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'O15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'O6'))))

def calc_INCOME_STATEMENT_V16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'V15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'V6'))))

def calc_INCOME_STATEMENT_AC16(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AC15')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AC6'))))

def calc_INCOME_STATEMENT_O10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'O11')), xl_ref(ctx.cell('INCOME STATEMENT', 'O13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'O17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'O15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'O19')))

def calc_INCOME_STATEMENT_O18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'O17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'O6'))))

def calc_INCOME_STATEMENT_AC18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AC17')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AC6'))))

def calc_INCOME_STATEMENT_K23(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'K21')), xl_ref(ctx.cell('INCOME STATEMENT', 'D21'))), 1), 100)

def calc_INCOME_STATEMENT_D25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'D21')), xl_ref(ctx.cell('INCOME STATEMENT', 'D6'))), 100)

def calc_INCOME_STATEMENT_D28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'D21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'D26')), xl_ref(ctx.cell('INCOME STATEMENT', 'D27'))))

def calc_INCOME_STATEMENT_F21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'F6')), xl_ref(ctx.cell('INCOME STATEMENT', 'F10')))

def calc_INCOME_STATEMENT_H20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'H19')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'H6'))))

def calc_INCOME_STATEMENT_M21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'M6')), xl_ref(ctx.cell('INCOME STATEMENT', 'M10')))

def calc_INCOME_STATEMENT_R23(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'R21')), xl_ref(ctx.cell('INCOME STATEMENT', 'K21'))), 1), 100)

def calc_INCOME_STATEMENT_R25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'R21')), xl_ref(ctx.cell('INCOME STATEMENT', 'R6'))), 100)

def calc_INCOME_STATEMENT_R28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'R21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'R26')), xl_ref(ctx.cell('INCOME STATEMENT', 'R27'))))

def calc_INCOME_STATEMENT_T21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'T6')), xl_ref(ctx.cell('INCOME STATEMENT', 'T10')))

def calc_INCOME_STATEMENT_V10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'V11')), xl_ref(ctx.cell('INCOME STATEMENT', 'V13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'V17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'V15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'V19')))

def calc_INCOME_STATEMENT_V20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'V19')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'V6'))))

def calc_INCOME_STATEMENT_Y23(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Y21')), xl_ref(ctx.cell('INCOME STATEMENT', 'R21'))), 1), 100)

def calc_INCOME_STATEMENT_Y25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Y21')), xl_ref(ctx.cell('INCOME STATEMENT', 'Y6'))), 100)

def calc_INCOME_STATEMENT_Y28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'Y21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'Y26')), xl_ref(ctx.cell('INCOME STATEMENT', 'Y27'))))

def calc_INCOME_STATEMENT_AA21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AA6')), xl_ref(ctx.cell('INCOME STATEMENT', 'AA10')))

def calc_INCOME_STATEMENT_AC10(ctx):
    return xl_add(xl_add(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AC11')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC13'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AC17'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AC15'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AC19')))

def calc_INCOME_STATEMENT_AC20(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AC19')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AC6'))))

def calc_COMPANY_OVERVIEW_G22(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AD21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AD6')))

def calc_INCOME_STATEMENT_AD25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AD21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AD6'))), 100)

def calc_INCOME_STATEMENT_AD28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AD21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AD26')), xl_ref(ctx.cell('INCOME STATEMENT', 'AD27'))))

def calc_Valuation_D20(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AD21'))

def calc_Ratio_Analysis_G13(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AD21')), 100), xl_ref(ctx.cell('INCOME STATEMENT', 'AD6')))

def calc_Ratio_Analysis_G19(ctx):
    return xl_div(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AD21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AD27'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AD26')))

def calc_Ratio_Analysis_G22(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AD21')), xl_ref(ctx.cell('INCOME STATEMENT', 'W21'))), 1)

def calc_COMPANY_OVERVIEW_H22(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AE21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AE6')))

def calc_INCOME_STATEMENT_AE23(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AE21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AD21'))), 1), 100)

def calc_INCOME_STATEMENT_AE25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AE21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AE6'))), 100)

def calc_INCOME_STATEMENT_AE28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AE21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AE26')), xl_ref(ctx.cell('INCOME STATEMENT', 'AE27'))))

def calc_Valuation_E20(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AE21'))

def calc_Ratio_Analysis_H13(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AE21')), 100), xl_ref(ctx.cell('INCOME STATEMENT', 'AE6')))

def calc_Ratio_Analysis_H19(ctx):
    return xl_div(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AE21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AE27'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AE26')))

def calc_Ratio_Analysis_H22(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AE21')), xl_ref(ctx.cell('INCOME STATEMENT', 'X21'))), 1)

def calc_COMPANY_OVERVIEW_I22(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AF21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AF6')))

def calc_INCOME_STATEMENT_AF23(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AF21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AE21'))), 1), 100)

def calc_INCOME_STATEMENT_AF25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AF21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AF6'))), 100)

def calc_INCOME_STATEMENT_AF28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AF21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AF26')), xl_ref(ctx.cell('INCOME STATEMENT', 'AF27'))))

def calc_Valuation_F20(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AF21'))

def calc_Ratio_Analysis_I13(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AF21')), 100), xl_ref(ctx.cell('INCOME STATEMENT', 'AF6')))

def calc_Ratio_Analysis_I19(ctx):
    return xl_div(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AF21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AF27'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AF26')))

def calc_Ratio_Analysis_I22(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AF21')), xl_ref(ctx.cell('INCOME STATEMENT', 'Y21'))), 1)

def calc_COMPANY_OVERVIEW_J22(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AG21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AG6')))

def calc_INCOME_STATEMENT_AG23(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AG21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AF21'))), 1), 100)

def calc_INCOME_STATEMENT_AG25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AG21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AG6'))), 100)

def calc_INCOME_STATEMENT_AG28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AG21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AG26')), xl_ref(ctx.cell('INCOME STATEMENT', 'AG27'))))

def calc_Valuation_G20(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AG21'))

def calc_Ratio_Analysis_J13(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AG21')), 100), xl_ref(ctx.cell('INCOME STATEMENT', 'AG6')))

def calc_Ratio_Analysis_J19(ctx):
    return xl_div(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AG21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AG27'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AG26')))

def calc_Ratio_Analysis_J22(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AG21')), xl_ref(ctx.cell('INCOME STATEMENT', 'Z21'))), 1)

def calc_Ratio_Analysis_C15(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'H26')), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'C19')), xl_ref(ctx.cell('BALANCESHEET', 'B19'))))

def calc_Ratio_Analysis_D15(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'O26')), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'D19')), xl_ref(ctx.cell('BALANCESHEET', 'C19'))))

def calc_Ratio_Analysis_E15(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'V26')), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'E19')), xl_ref(ctx.cell('BALANCESHEET', 'D19'))))

def calc_Ratio_Analysis_F15(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AC26')), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'F19')), xl_ref(ctx.cell('BALANCESHEET', 'E19'))))

def calc_CASH_FOW_STATEMENT_B11(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'H27'))

def calc_CASH_FOW_STATEMENT_C11(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'O27'))

def calc_INCOME_STATEMENT_K32(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'K30')), xl_ref(ctx.cell('INCOME STATEMENT', 'K31')))

def calc_CASH_FOW_STATEMENT_D11(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'V27'))

def calc_Valuation_B24(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'V27'))

def calc_Valuation_C22(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AC33'))

def calc_CASH_FOW_STATEMENT_B9(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'H33')), xl_ref(ctx.cell('INCOME STATEMENT', 'H34')))

def calc_CASH_FOW_STATEMENT_E9(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AC33')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC34')))

def calc_INCOME_STATEMENT_F47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'E47'))

def calc_INCOME_STATEMENT_E48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'E47')), 10.52)

def calc_INCOME_STATEMENT_W47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'V47'))

def calc_INCOME_STATEMENT_V48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'V47')), 10)

def calc_CASH_FOW_STATEMENT_I14(ctx):
    return xl_sub(xl_ref(ctx.cell('CASH FOW STATEMENT', 'I12')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'I13')))

def calc_Valuation_F46(ctx):
    return xl_ref(ctx.cell('Valuation', 'E46'))

def calc_Valuation_I25(ctx):
    return xl_ref(ctx.cell('Valuation', 'H25'))

def calc_Valuation_G13(ctx):
    return xl_add(xl_ref(ctx.cell('Valuation', 'G9')), xl_ref(ctx.cell('Valuation', 'G12')))

def calc_Valuation_G55(ctx):
    return xl_add(xl_ref(ctx.cell('Valuation', 'G53')), xl_ref(ctx.cell('Valuation', 'G54')))

def calc_Valuation_B25(ctx):
    return xl_ref(ctx.cell('CASH FOW STATEMENT', 'D20'))

def calc_Valuation_C25(ctx):
    return xl_ref(ctx.cell('CASH FOW STATEMENT', 'E20'))

def calc_Valuation_G37(ctx):
    return xl_sub(xl_ref(ctx.cell('Valuation', 'F37')), xl_div(xl_sub(xl_ref(ctx.cell('Valuation', 'E37')), xl_ref(ctx.cell('Valuation', 'J37'))), 5))

def calc_Valuation_G45(ctx):
    return xl_ref(ctx.cell('Valuation', 'F45'))

def calc_PRESENTATION_B40(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'B38')), xl_ref(ctx.cell('PRESENTATION', 'B39')))

def calc_PRESENTATION_C40(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'C38')), xl_ref(ctx.cell('PRESENTATION', 'C39')))

def calc_PRESENTATION_E40(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'E38')), xl_ref(ctx.cell('PRESENTATION', 'E39')))

def calc_PRESENTATION_G40(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'G38')), xl_ref(ctx.cell('PRESENTATION', 'G39')))

def calc_PRESENTATION_I40(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'I38')), xl_ref(ctx.cell('PRESENTATION', 'I39')))

def calc_PRESENTATION_J40(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'J38')), xl_ref(ctx.cell('PRESENTATION', 'J39')))

def calc_PRESENTATION_L40(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'L38')), xl_ref(ctx.cell('PRESENTATION', 'L39')))

def calc_PRESENTATION_N40(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'N38')), xl_ref(ctx.cell('PRESENTATION', 'N39')))

def calc_PRESENTATION_P40(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'P38')), xl_ref(ctx.cell('PRESENTATION', 'P39')))

def calc_PRESENTATION_Q40(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'Q38')), xl_ref(ctx.cell('PRESENTATION', 'Q39')))

def calc_PRESENTATION_S40(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'S38')), xl_ref(ctx.cell('PRESENTATION', 'S39')))

def calc_PRESENTATION_U40(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'U38')), xl_ref(ctx.cell('PRESENTATION', 'U39')))

def calc_PRESENTATION_W40(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'W38')), xl_ref(ctx.cell('PRESENTATION', 'W39')))

def calc_PRESENTATION_X40(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'X38')), xl_ref(ctx.cell('PRESENTATION', 'X39')))

def calc_PRESENTATION_Z40(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'Z38')), xl_ref(ctx.cell('PRESENTATION', 'Z39')))

def calc_PRESENTATION_AB40(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AB38')), xl_ref(ctx.cell('PRESENTATION', 'AB39')))

def calc_PRESENTATION_AD40(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AD38')), xl_ref(ctx.cell('PRESENTATION', 'AD39')))

def calc_PRESENTATION_AE40(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AE38')), xl_ref(ctx.cell('PRESENTATION', 'AE39')))

def calc_PRESENTATION_AF40(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AF38')), xl_ref(ctx.cell('PRESENTATION', 'AF39')))

def calc_PRESENTATION_AG40(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AG38')), xl_ref(ctx.cell('PRESENTATION', 'AG39')))

def calc_PRESENTATION_H32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'H25')), xl_ref(ctx.cell('PRESENTATION', 'H26')))

def calc_PRESENTATION_O32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'O25')), xl_ref(ctx.cell('PRESENTATION', 'O26')))

def calc_PRESENTATION_D38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'D36')), xl_ref(ctx.cell('PRESENTATION', 'D37')))

def calc_PRESENTATION_F33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'F32')), xl_ref(ctx.cell('PRESENTATION', 'F25'))), 100)

def calc_PRESENTATION_F36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'F32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'F34')), xl_ref(ctx.cell('PRESENTATION', 'F35'))))

def calc_PRESENTATION_M33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'M32')), xl_ref(ctx.cell('PRESENTATION', 'M25'))), 100)

def calc_PRESENTATION_M36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'M32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'M34')), xl_ref(ctx.cell('PRESENTATION', 'M35'))))

def calc_PRESENTATION_R38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'R36')), xl_ref(ctx.cell('PRESENTATION', 'R37')))

def calc_PRESENTATION_T33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'T32')), xl_ref(ctx.cell('PRESENTATION', 'T25'))), 100)

def calc_PRESENTATION_T36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'T32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'T34')), xl_ref(ctx.cell('PRESENTATION', 'T35'))))

def calc_PRESENTATION_V32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'V25')), xl_ref(ctx.cell('PRESENTATION', 'V26')))

def calc_PRESENTATION_Y38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'Y36')), xl_ref(ctx.cell('PRESENTATION', 'Y37')))

def calc_PRESENTATION_AA33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'AA32')), xl_ref(ctx.cell('PRESENTATION', 'AA25'))), 100)

def calc_PRESENTATION_AA36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AA32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'AA34')), xl_ref(ctx.cell('PRESENTATION', 'AA35'))))

def calc_PRESENTATION_AC32(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AC25')), xl_ref(ctx.cell('PRESENTATION', 'AC26')))

def calc_PRESENTATION_K44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'K40')), xl_ref(ctx.cell('PRESENTATION', 'K41'))), xl_ref(ctx.cell('PRESENTATION', 'K42'))), xl_ref(ctx.cell('PRESENTATION', 'K43')))

def calc_Segment_Revenue_Model_X8(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'X24')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'W8')))

def calc_Segment_Revenue_Model_X9(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'X25')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'W9')))

def calc_Segment_Revenue_Model_X10(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'X26')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'W10')))

def calc_Segment_Revenue_Model_X11(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'X27')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'W11')))

def calc_Segment_Revenue_Model_Y12(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'Y28')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'X12')))

def calc_Segment_Revenue_Model_X13(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'X29')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'W13')))

def calc_Segment_Revenue_Model_U15(ctx):
    return xl_add(xl_add(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'Q15')), xl_ref(ctx.cell('Segment Revenue Model', 'R15'))), xl_ref(ctx.cell('Segment Revenue Model', 'S15'))), xl_ref(ctx.cell('Segment Revenue Model', 'T15')))

def calc_Segment_Revenue_Model_T17(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'T15')), xl_ref(ctx.cell('Segment Revenue Model', 'T16')))

def calc_Segment_Revenue_Model_T31(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'T15')), xl_ref(ctx.cell('Segment Revenue Model', 'O15'))), 1)

def calc_Segment_Revenue_Model_X14(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'X30')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'W14')))

def calc_Segment_Revenue_Model_W15(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'W31')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'V15')))

def calc_Segment_Revenue_Model_V17(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'V15')), xl_ref(ctx.cell('Segment Revenue Model', 'V16')))

def calc_Segment_Revenue_Model_Y16(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'Y32')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'X16')))

def calc_INCOME_STATEMENT_O7(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'M7')), xl_ref(ctx.cell('INCOME STATEMENT', 'N7')))

def calc_INCOME_STATEMENT_AC7(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AA7')), xl_ref(ctx.cell('INCOME STATEMENT', 'AB7')))

def calc_Valuation_C37(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'C19')), xl_ref(ctx.cell('Valuation', 'B19'))), 1)

def calc_Valuation_D37(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'D19')), xl_ref(ctx.cell('Valuation', 'C19'))), 1)

def calc_Valuation_C45(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'C24')), xl_ref(ctx.cell('Valuation', 'C19')))

def calc_Valuation_C46(ctx):
    return xl_div(xl_uminus(xl_ref(ctx.cell('Valuation', 'C27'))), xl_ref(ctx.cell('Valuation', 'C19')))

def calc_INCOME_STATEMENT_B32(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'B30')), xl_ref(ctx.cell('INCOME STATEMENT', 'B31')))

def calc_INCOME_STATEMENT_C32(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'C30')), xl_ref(ctx.cell('INCOME STATEMENT', 'C31')))

def calc_INCOME_STATEMENT_E32(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'E30')), xl_ref(ctx.cell('INCOME STATEMENT', 'E31')))

def calc_INCOME_STATEMENT_G32(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'G30')), xl_ref(ctx.cell('INCOME STATEMENT', 'G51')))

def calc_INCOME_STATEMENT_I32(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'I30')), xl_ref(ctx.cell('INCOME STATEMENT', 'I31')))

def calc_INCOME_STATEMENT_J32(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'J30')), xl_ref(ctx.cell('INCOME STATEMENT', 'J31')))

def calc_INCOME_STATEMENT_L32(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'L30')), xl_ref(ctx.cell('INCOME STATEMENT', 'L31')))

def calc_INCOME_STATEMENT_N32(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'N30')), xl_ref(ctx.cell('INCOME STATEMENT', 'N31')))

def calc_INCOME_STATEMENT_P32(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'P30')), xl_ref(ctx.cell('INCOME STATEMENT', 'P31')))

def calc_INCOME_STATEMENT_Q32(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'Q30')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q31')))

def calc_INCOME_STATEMENT_S32(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'S30')), xl_ref(ctx.cell('INCOME STATEMENT', 'S31')))

def calc_INCOME_STATEMENT_U32(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'U30')), xl_ref(ctx.cell('INCOME STATEMENT', 'U31')))

def calc_INCOME_STATEMENT_W32(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'W30')), xl_ref(ctx.cell('INCOME STATEMENT', 'W31')))

def calc_INCOME_STATEMENT_X32(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'X30')), xl_ref(ctx.cell('INCOME STATEMENT', 'X31')))

def calc_INCOME_STATEMENT_Z32(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'Z30')), xl_ref(ctx.cell('INCOME STATEMENT', 'Z31')))

def calc_INCOME_STATEMENT_AB32(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AB30')), xl_ref(ctx.cell('INCOME STATEMENT', 'AB31')))

def calc_PRESENTATION_C64(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'H6')), xl_ref(ctx.cell('INCOME STATEMENT', 'H10')))

def calc_INCOME_STATEMENT_H21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'H6')), xl_ref(ctx.cell('INCOME STATEMENT', 'H10')))

def calc_PRESENTATION_D64(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'O6')), xl_ref(ctx.cell('INCOME STATEMENT', 'O10')))

def calc_INCOME_STATEMENT_O21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'O6')), xl_ref(ctx.cell('INCOME STATEMENT', 'O10')))

def calc_INCOME_STATEMENT_D30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'D28')), xl_ref(ctx.cell('INCOME STATEMENT', 'D29')))

def calc_INCOME_STATEMENT_F25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'F21')), xl_ref(ctx.cell('INCOME STATEMENT', 'F6'))), 100)

def calc_INCOME_STATEMENT_F28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'F21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'F26')), xl_ref(ctx.cell('INCOME STATEMENT', 'F27'))))

def calc_INCOME_STATEMENT_M23(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'M21')), xl_ref(ctx.cell('INCOME STATEMENT', 'F21'))), 1), 100)

def calc_INCOME_STATEMENT_M25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'M21')), xl_ref(ctx.cell('INCOME STATEMENT', 'M6'))), 100)

def calc_INCOME_STATEMENT_M28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'M21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'M26')), xl_ref(ctx.cell('INCOME STATEMENT', 'M27'))))

def calc_INCOME_STATEMENT_R30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'R28')), xl_ref(ctx.cell('INCOME STATEMENT', 'R29')))

def calc_INCOME_STATEMENT_T23(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'T21')), xl_ref(ctx.cell('INCOME STATEMENT', 'M21'))), 1), 100)

def calc_INCOME_STATEMENT_T25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'T21')), xl_ref(ctx.cell('INCOME STATEMENT', 'T6'))), 100)

def calc_INCOME_STATEMENT_T28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'T21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'T26')), xl_ref(ctx.cell('INCOME STATEMENT', 'T27'))))

def calc_PRESENTATION_E64(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'V6')), xl_ref(ctx.cell('INCOME STATEMENT', 'V10')))

def calc_INCOME_STATEMENT_V21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'V6')), xl_ref(ctx.cell('INCOME STATEMENT', 'V10')))

def calc_INCOME_STATEMENT_Y30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'Y28')), xl_ref(ctx.cell('INCOME STATEMENT', 'Y29')))

def calc_INCOME_STATEMENT_AA23(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AA21')), xl_ref(ctx.cell('INCOME STATEMENT', 'T21'))), 1), 100)

def calc_INCOME_STATEMENT_AA25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AA21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AA6'))), 100)

def calc_INCOME_STATEMENT_AA28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AA21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AA26')), xl_ref(ctx.cell('INCOME STATEMENT', 'AA27'))))

def calc_PRESENTATION_F64(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AC6')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC10')))

def calc_INCOME_STATEMENT_AC21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AC6')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC10')))

def calc_INCOME_STATEMENT_AD30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AD28')), xl_ref(ctx.cell('INCOME STATEMENT', 'AD29')))

def calc_Valuation_D21(ctx):
    return xl_sub(xl_ref(ctx.cell('Valuation', 'D20')), xl_ref(ctx.cell('INCOME STATEMENT', 'AD27')))

def calc_Valuation_D43(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'D20')), xl_ref(ctx.cell('Valuation', 'D19')))

def calc_PRESENTATION_G63(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'G13'))

def calc_COMPANY_OVERVIEW_G24(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'G19'))

def calc_PRESENTATION_G66(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'G19'))

def calc_INCOME_STATEMENT_AE30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AE28')), xl_ref(ctx.cell('INCOME STATEMENT', 'AE29')))

def calc_Valuation_E21(ctx):
    return xl_sub(xl_ref(ctx.cell('Valuation', 'E20')), xl_ref(ctx.cell('INCOME STATEMENT', 'AE27')))

def calc_Valuation_E38(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'E20')), xl_ref(ctx.cell('Valuation', 'D20'))), 1)

def calc_Valuation_E43(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'E20')), xl_ref(ctx.cell('Valuation', 'E19')))

def calc_PRESENTATION_H63(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'H13'))

def calc_COMPANY_OVERVIEW_H24(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'H19'))

def calc_PRESENTATION_H66(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'H19'))

def calc_INCOME_STATEMENT_AF30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AF28')), xl_ref(ctx.cell('INCOME STATEMENT', 'AF29')))

def calc_Valuation_F21(ctx):
    return xl_sub(xl_ref(ctx.cell('Valuation', 'F20')), xl_ref(ctx.cell('INCOME STATEMENT', 'AF27')))

def calc_Valuation_F38(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'F20')), xl_ref(ctx.cell('Valuation', 'E20'))), 1)

def calc_PRESENTATION_I63(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'I13'))

def calc_COMPANY_OVERVIEW_I24(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'I19'))

def calc_PRESENTATION_I66(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'I19'))

def calc_INCOME_STATEMENT_AG30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AG28')), xl_ref(ctx.cell('INCOME STATEMENT', 'AG29')))

def calc_Valuation_G21(ctx):
    return xl_sub(xl_ref(ctx.cell('Valuation', 'G20')), xl_ref(ctx.cell('INCOME STATEMENT', 'AG27')))

def calc_Valuation_G38(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'G20')), xl_ref(ctx.cell('Valuation', 'F20'))), 1)

def calc_PRESENTATION_J63(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'J13'))

def calc_COMPANY_OVERVIEW_J24(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'J19'))

def calc_PRESENTATION_J66(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'J19'))

def calc_INCOME_STATEMENT_K36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'K33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'K32'))))

def calc_INCOME_STATEMENT_K37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'K33')), xl_ref(ctx.cell('INCOME STATEMENT', 'K34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'K35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'K32'))), 100)

def calc_INCOME_STATEMENT_K38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'K32')), xl_ref(ctx.cell('INCOME STATEMENT', 'K33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'K34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'K35')))

def calc_Valuation_C40(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'C24')), xl_ref(ctx.cell('Valuation', 'B24'))), 1)

def calc_Valuation_B45(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'B24')), xl_ref(ctx.cell('Valuation', 'B19')))

def calc_INCOME_STATEMENT_G47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'F47'))

def calc_INCOME_STATEMENT_F48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'F47')), 10)

def calc_INCOME_STATEMENT_X47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'W47'))

def calc_INCOME_STATEMENT_W48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'W47')), 10)

def calc_Valuation_B30(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'V48'))

def calc_CASH_FOW_STATEMENT_I23(ctx):
    return xl_add(xl_add(xl_sub(xl_sub(xl_add(xl_sub(xl_ref(ctx.cell('CASH FOW STATEMENT', 'I14')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'I16'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'I17'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'I19'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'H20'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'I18'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'I21')))

def calc_Valuation_G46(ctx):
    return xl_ref(ctx.cell('Valuation', 'F46'))

def calc_Valuation_J25(ctx):
    return xl_ref(ctx.cell('Valuation', 'I25'))

def calc_COMPANY_OVERVIEW_E25(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'G13')), xl_ref(ctx.cell('INCOME STATEMENT', 'V6')))

def calc_COMPANY_OVERVIEW_F25(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'G13')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC6')))

def calc_COMPANY_OVERVIEW_G25(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'G13')), xl_ref(ctx.cell('INCOME STATEMENT', 'AD6')))

def calc_COMPANY_OVERVIEW_H25(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'G13')), xl_ref(ctx.cell('INCOME STATEMENT', 'AE6')))

def calc_COMPANY_OVERVIEW_I25(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'G13')), xl_ref(ctx.cell('INCOME STATEMENT', 'AF6')))

def calc_COMPANY_OVERVIEW_J25(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'G13')), xl_ref(ctx.cell('INCOME STATEMENT', 'AG6')))

def calc_COMPANY_OVERVIEW_G26(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'G13')), xl_ref(ctx.cell('INCOME STATEMENT', 'AD21')))

def calc_COMPANY_OVERVIEW_H26(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'G13')), xl_ref(ctx.cell('INCOME STATEMENT', 'AE21')))

def calc_COMPANY_OVERVIEW_I26(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'G13')), xl_ref(ctx.cell('INCOME STATEMENT', 'AF21')))

def calc_COMPANY_OVERVIEW_J26(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'G13')), xl_ref(ctx.cell('INCOME STATEMENT', 'AG21')))

def calc_Valuation_K6(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'G9')), xl_ref(ctx.cell('Valuation', 'G13')))

def calc_Valuation_K7(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'G12')), xl_ref(ctx.cell('Valuation', 'G13')))

def calc_Valuation_H37(ctx):
    return xl_sub(xl_ref(ctx.cell('Valuation', 'G37')), xl_div(xl_sub(xl_ref(ctx.cell('Valuation', 'E37')), xl_ref(ctx.cell('Valuation', 'J37'))), 5))

def calc_Valuation_H45(ctx):
    return xl_ref(ctx.cell('Valuation', 'G45'))

def calc_PRESENTATION_B44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'B40')), xl_ref(ctx.cell('PRESENTATION', 'B41'))), xl_ref(ctx.cell('PRESENTATION', 'B42'))), xl_ref(ctx.cell('PRESENTATION', 'B43')))

def calc_PRESENTATION_C44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'C40')), xl_ref(ctx.cell('PRESENTATION', 'C41'))), xl_ref(ctx.cell('PRESENTATION', 'C42'))), xl_ref(ctx.cell('PRESENTATION', 'C43')))

def calc_PRESENTATION_E44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'E40')), xl_ref(ctx.cell('PRESENTATION', 'E41'))), xl_ref(ctx.cell('PRESENTATION', 'E42'))), xl_ref(ctx.cell('PRESENTATION', 'E43')))

def calc_PRESENTATION_G44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'G40')), xl_ref(ctx.cell('PRESENTATION', 'G41'))), xl_ref(ctx.cell('PRESENTATION', 'G42'))), xl_ref(ctx.cell('PRESENTATION', 'G43')))

def calc_PRESENTATION_I44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'I40')), xl_ref(ctx.cell('PRESENTATION', 'I41'))), xl_ref(ctx.cell('PRESENTATION', 'I42'))), xl_ref(ctx.cell('PRESENTATION', 'I43')))

def calc_PRESENTATION_J44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'J40')), xl_ref(ctx.cell('PRESENTATION', 'J41'))), xl_ref(ctx.cell('PRESENTATION', 'J42'))), xl_ref(ctx.cell('PRESENTATION', 'J43')))

def calc_PRESENTATION_L44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'L40')), xl_ref(ctx.cell('PRESENTATION', 'L41'))), xl_ref(ctx.cell('PRESENTATION', 'L42'))), xl_ref(ctx.cell('PRESENTATION', 'L43')))

def calc_PRESENTATION_N44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'N40')), xl_ref(ctx.cell('PRESENTATION', 'N41'))), xl_ref(ctx.cell('PRESENTATION', 'N42'))), xl_ref(ctx.cell('PRESENTATION', 'N43')))

def calc_PRESENTATION_P44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'P40')), xl_ref(ctx.cell('PRESENTATION', 'P41'))), xl_ref(ctx.cell('PRESENTATION', 'P42'))), xl_ref(ctx.cell('PRESENTATION', 'P43')))

def calc_PRESENTATION_Q44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'Q40')), xl_ref(ctx.cell('PRESENTATION', 'Q41'))), xl_ref(ctx.cell('PRESENTATION', 'Q42'))), xl_ref(ctx.cell('PRESENTATION', 'Q43')))

def calc_PRESENTATION_S44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'S40')), xl_ref(ctx.cell('PRESENTATION', 'S41'))), xl_ref(ctx.cell('PRESENTATION', 'S42'))), xl_ref(ctx.cell('PRESENTATION', 'S43')))

def calc_PRESENTATION_U44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'U40')), xl_ref(ctx.cell('PRESENTATION', 'U41'))), xl_ref(ctx.cell('PRESENTATION', 'U42'))), xl_ref(ctx.cell('PRESENTATION', 'U43')))

def calc_PRESENTATION_W44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'W40')), xl_ref(ctx.cell('PRESENTATION', 'W41'))), xl_ref(ctx.cell('PRESENTATION', 'W42'))), xl_ref(ctx.cell('PRESENTATION', 'W43')))

def calc_PRESENTATION_X44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'X40')), xl_ref(ctx.cell('PRESENTATION', 'X41'))), xl_ref(ctx.cell('PRESENTATION', 'X42'))), xl_ref(ctx.cell('PRESENTATION', 'X43')))

def calc_PRESENTATION_Z44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'Z40')), xl_ref(ctx.cell('PRESENTATION', 'Z41'))), xl_ref(ctx.cell('PRESENTATION', 'Z42'))), xl_ref(ctx.cell('PRESENTATION', 'Z43')))

def calc_PRESENTATION_AB44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AB40')), xl_ref(ctx.cell('PRESENTATION', 'AB41'))), xl_ref(ctx.cell('PRESENTATION', 'AB42'))), xl_ref(ctx.cell('PRESENTATION', 'AB43')))

def calc_PRESENTATION_AD41(ctx):
    return xl_mul(xl_ref(ctx.cell('PRESENTATION', 'AD40')), xl_percent(33))

def calc_PRESENTATION_AE41(ctx):
    return xl_mul(xl_ref(ctx.cell('PRESENTATION', 'AE40')), xl_percent(33))

def calc_PRESENTATION_AF41(ctx):
    return xl_mul(xl_ref(ctx.cell('PRESENTATION', 'AF40')), xl_percent(33))

def calc_PRESENTATION_AG41(ctx):
    return xl_mul(xl_ref(ctx.cell('PRESENTATION', 'AG40')), xl_percent(33))

def calc_PRESENTATION_H33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'H32')), xl_ref(ctx.cell('PRESENTATION', 'H25'))), 100)

def calc_PRESENTATION_H36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'H32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'H34')), xl_ref(ctx.cell('PRESENTATION', 'H35'))))

def calc_PRESENTATION_O33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'O32')), xl_ref(ctx.cell('PRESENTATION', 'O25'))), 100)

def calc_PRESENTATION_O36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'O32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'O34')), xl_ref(ctx.cell('PRESENTATION', 'O35'))))

def calc_PRESENTATION_D40(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'D38')), xl_ref(ctx.cell('PRESENTATION', 'D39')))

def calc_PRESENTATION_F38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'F36')), xl_ref(ctx.cell('PRESENTATION', 'F37')))

def calc_PRESENTATION_M38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'M36')), xl_ref(ctx.cell('PRESENTATION', 'M37')))

def calc_PRESENTATION_R40(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'R38')), xl_ref(ctx.cell('PRESENTATION', 'R39')))

def calc_PRESENTATION_T38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'T36')), xl_ref(ctx.cell('PRESENTATION', 'T37')))

def calc_PRESENTATION_V33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'V32')), xl_ref(ctx.cell('PRESENTATION', 'V25'))), 100)

def calc_PRESENTATION_V36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'V32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'V34')), xl_ref(ctx.cell('PRESENTATION', 'V35'))))

def calc_PRESENTATION_Y40(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'Y38')), xl_ref(ctx.cell('PRESENTATION', 'Y39')))

def calc_PRESENTATION_AA38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'AA36')), xl_ref(ctx.cell('PRESENTATION', 'AA37')))

def calc_PRESENTATION_AC33(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('PRESENTATION', 'AC32')), xl_ref(ctx.cell('PRESENTATION', 'AC25'))), 100)

def calc_PRESENTATION_AC36(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AC32')), xl_add(xl_ref(ctx.cell('PRESENTATION', 'AC34')), xl_ref(ctx.cell('PRESENTATION', 'AC35'))))

def calc_PRESENTATION_K45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'K44'))

def calc_PRESENTATION_K47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'K37')), xl_ref(ctx.cell('PRESENTATION', 'K44')))

def calc_Segment_Revenue_Model_Y8(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'Y24')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'X8')))

def calc_Segment_Revenue_Model_Y9(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'Y25')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'X9')))

def calc_Segment_Revenue_Model_Y10(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'Y26')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'X10')))

def calc_Segment_Revenue_Model_Y11(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'Y27')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'X11')))

def calc_Segment_Revenue_Model_Y13(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'Y29')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'X13')))

def calc_Segment_Revenue_Model_U17(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'U15')), xl_ref(ctx.cell('Segment Revenue Model', 'U16')))

def calc_Segment_Revenue_Model_T33(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'T17')), xl_ref(ctx.cell('Segment Revenue Model', 'O17'))), 1)

def calc_Segment_Revenue_Model_T37(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'T8')), xl_ref(ctx.cell('Segment Revenue Model', 'T17')))

def calc_Segment_Revenue_Model_T38(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'T9')), xl_ref(ctx.cell('Segment Revenue Model', 'T17')))

def calc_Segment_Revenue_Model_T39(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'T10')), xl_ref(ctx.cell('Segment Revenue Model', 'T17')))

def calc_Segment_Revenue_Model_T40(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'T11')), xl_ref(ctx.cell('Segment Revenue Model', 'T17')))

def calc_Segment_Revenue_Model_T41(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'T12')), xl_ref(ctx.cell('Segment Revenue Model', 'T17')))

def calc_Segment_Revenue_Model_T42(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'T13')), xl_ref(ctx.cell('Segment Revenue Model', 'T17')))

def calc_Segment_Revenue_Model_T43(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'T14')), xl_ref(ctx.cell('Segment Revenue Model', 'T17')))

def calc_Segment_Revenue_Model_T44(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'T15')), xl_ref(ctx.cell('Segment Revenue Model', 'T17')))

def calc_Segment_Revenue_Model_T45(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'T16')), xl_ref(ctx.cell('Segment Revenue Model', 'T17')))

def calc_Segment_Revenue_Model_T46(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'T17')), xl_ref(ctx.cell('Segment Revenue Model', 'T17')))

def calc_Segment_Revenue_Model_Y14(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'Y30')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'X14')))

def calc_Segment_Revenue_Model_X15(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'X31')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'W15')))

def calc_Segment_Revenue_Model_W17(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'W15')), xl_ref(ctx.cell('Segment Revenue Model', 'W16')))

def calc_Segment_Revenue_Model_V37(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'V8')), xl_ref(ctx.cell('Segment Revenue Model', 'V17')))

def calc_Segment_Revenue_Model_V38(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'V9')), xl_ref(ctx.cell('Segment Revenue Model', 'V17')))

def calc_Segment_Revenue_Model_V39(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'V10')), xl_ref(ctx.cell('Segment Revenue Model', 'V17')))

def calc_Segment_Revenue_Model_V40(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'V11')), xl_ref(ctx.cell('Segment Revenue Model', 'V17')))

def calc_Segment_Revenue_Model_V41(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'V12')), xl_ref(ctx.cell('Segment Revenue Model', 'V17')))

def calc_Segment_Revenue_Model_V42(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'V13')), xl_ref(ctx.cell('Segment Revenue Model', 'V17')))

def calc_Segment_Revenue_Model_V43(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'V14')), xl_ref(ctx.cell('Segment Revenue Model', 'V17')))

def calc_Segment_Revenue_Model_V44(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'V15')), xl_ref(ctx.cell('Segment Revenue Model', 'V17')))

def calc_Segment_Revenue_Model_V45(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'V16')), xl_ref(ctx.cell('Segment Revenue Model', 'V17')))

def calc_Segment_Revenue_Model_V46(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'V17')), xl_ref(ctx.cell('Segment Revenue Model', 'V17')))

def calc_INCOME_STATEMENT_B36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'B33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'B32'))))

def calc_INCOME_STATEMENT_B37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'B33')), xl_ref(ctx.cell('INCOME STATEMENT', 'B34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'B35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'B32'))), 100)

def calc_INCOME_STATEMENT_B38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'B32')), xl_ref(ctx.cell('INCOME STATEMENT', 'B33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'B34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'B35')))

def calc_INCOME_STATEMENT_C36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'C33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'C32'))))

def calc_INCOME_STATEMENT_C37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'C33')), xl_ref(ctx.cell('INCOME STATEMENT', 'C34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'C35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'C32'))), 100)

def calc_INCOME_STATEMENT_C38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'C32')), xl_ref(ctx.cell('INCOME STATEMENT', 'C33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'C34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'C35')))

def calc_INCOME_STATEMENT_E36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'E33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'E32'))))

def calc_INCOME_STATEMENT_E37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'E33')), xl_ref(ctx.cell('INCOME STATEMENT', 'E34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'E35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'E32'))), 100)

def calc_INCOME_STATEMENT_E38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'E32')), xl_ref(ctx.cell('INCOME STATEMENT', 'E33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'E34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'E35')))

def calc_INCOME_STATEMENT_G36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'G33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'G32'))))

def calc_INCOME_STATEMENT_G37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'G33')), xl_ref(ctx.cell('INCOME STATEMENT', 'G34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'G35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'G32'))), 100)

def calc_INCOME_STATEMENT_G38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'G32')), xl_ref(ctx.cell('INCOME STATEMENT', 'G33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'G34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'G35')))

def calc_INCOME_STATEMENT_I36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'I33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'I32'))))

def calc_INCOME_STATEMENT_I37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'I33')), xl_ref(ctx.cell('INCOME STATEMENT', 'I34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'I35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'I32'))), 100)

def calc_INCOME_STATEMENT_I38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'I32')), xl_ref(ctx.cell('INCOME STATEMENT', 'I33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'I34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'I35')))

def calc_INCOME_STATEMENT_J36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'J33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'J32'))))

def calc_INCOME_STATEMENT_J37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'J33')), xl_ref(ctx.cell('INCOME STATEMENT', 'J34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'J35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'J32'))), 100)

def calc_INCOME_STATEMENT_J38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'J32')), xl_ref(ctx.cell('INCOME STATEMENT', 'J33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'J34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'J35')))

def calc_INCOME_STATEMENT_L36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'L33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'L32'))))

def calc_INCOME_STATEMENT_L37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'L33')), xl_ref(ctx.cell('INCOME STATEMENT', 'L34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'L35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'L32'))), 100)

def calc_INCOME_STATEMENT_L38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'L32')), xl_ref(ctx.cell('INCOME STATEMENT', 'L33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'L34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'L35')))

def calc_INCOME_STATEMENT_N36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'N33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'N32'))))

def calc_INCOME_STATEMENT_N37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'N33')), xl_ref(ctx.cell('INCOME STATEMENT', 'N34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'N35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'N32'))), 100)

def calc_INCOME_STATEMENT_N38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'N32')), xl_ref(ctx.cell('INCOME STATEMENT', 'N33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'N34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'N35')))

def calc_INCOME_STATEMENT_P36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'P33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'P32'))))

def calc_INCOME_STATEMENT_P37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'P33')), xl_ref(ctx.cell('INCOME STATEMENT', 'P34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'P35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'P32'))), 100)

def calc_INCOME_STATEMENT_P38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'P32')), xl_ref(ctx.cell('INCOME STATEMENT', 'P33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'P34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'P35')))

def calc_INCOME_STATEMENT_Q36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Q33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'Q32'))))

def calc_INCOME_STATEMENT_Q37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'Q33')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'Q35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'Q32'))), 100)

def calc_INCOME_STATEMENT_Q38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'Q32')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'Q34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'Q35')))

def calc_INCOME_STATEMENT_S36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'S33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'S32'))))

def calc_INCOME_STATEMENT_S37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'S33')), xl_ref(ctx.cell('INCOME STATEMENT', 'S34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'S35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'S32'))), 100)

def calc_INCOME_STATEMENT_S38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'S32')), xl_ref(ctx.cell('INCOME STATEMENT', 'S33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'S34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'S35')))

def calc_INCOME_STATEMENT_U36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'U33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'U32'))))

def calc_INCOME_STATEMENT_U37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'U33')), xl_ref(ctx.cell('INCOME STATEMENT', 'U34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'U35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'U32'))), 100)

def calc_INCOME_STATEMENT_U38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'U32')), xl_ref(ctx.cell('INCOME STATEMENT', 'U33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'U34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'U35')))

def calc_INCOME_STATEMENT_W36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'W33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'W32'))))

def calc_INCOME_STATEMENT_W37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'W33')), xl_ref(ctx.cell('INCOME STATEMENT', 'W34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'W35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'W32'))), 100)

def calc_INCOME_STATEMENT_W38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'W32')), xl_ref(ctx.cell('INCOME STATEMENT', 'W33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'W34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'W35')))

def calc_INCOME_STATEMENT_X36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'X33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'X32'))))

def calc_INCOME_STATEMENT_X37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'X33')), xl_ref(ctx.cell('INCOME STATEMENT', 'X34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'X35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'X32'))), 100)

def calc_INCOME_STATEMENT_X38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'X32')), xl_ref(ctx.cell('INCOME STATEMENT', 'X33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'X34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'X35')))

def calc_INCOME_STATEMENT_Z36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Z33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'Z32'))))

def calc_INCOME_STATEMENT_Z37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'Z33')), xl_ref(ctx.cell('INCOME STATEMENT', 'Z34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'Z35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'Z32'))), 100)

def calc_INCOME_STATEMENT_Z38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'Z32')), xl_ref(ctx.cell('INCOME STATEMENT', 'Z33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'Z34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'Z35')))

def calc_INCOME_STATEMENT_AB36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AB33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AB32'))))

def calc_INCOME_STATEMENT_AB37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AB33')), xl_ref(ctx.cell('INCOME STATEMENT', 'AB34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AB35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AB32'))), 100)

def calc_INCOME_STATEMENT_AB38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AB32')), xl_ref(ctx.cell('INCOME STATEMENT', 'AB33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AB34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AB35')))

def calc_COMPANY_OVERVIEW_C22(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'H21')), xl_ref(ctx.cell('INCOME STATEMENT', 'H6')))

def calc_INCOME_STATEMENT_B22(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'B21')), xl_ref(ctx.cell('INCOME STATEMENT', 'H21'))), 100)

def calc_INCOME_STATEMENT_C22(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'C21')), xl_ref(ctx.cell('INCOME STATEMENT', 'H21'))), 100)

def calc_INCOME_STATEMENT_D22(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'D21')), xl_ref(ctx.cell('INCOME STATEMENT', 'H21'))), 100)

def calc_INCOME_STATEMENT_E22(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'E21')), xl_ref(ctx.cell('INCOME STATEMENT', 'H21'))), 100)

def calc_INCOME_STATEMENT_F22(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'F21')), xl_ref(ctx.cell('INCOME STATEMENT', 'H21'))), 100)

def calc_INCOME_STATEMENT_G22(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'G21')), xl_ref(ctx.cell('INCOME STATEMENT', 'H21'))), 100)

def calc_INCOME_STATEMENT_H25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'H21')), xl_ref(ctx.cell('INCOME STATEMENT', 'H6'))), 100)

def calc_INCOME_STATEMENT_H28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'H21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'H26')), xl_ref(ctx.cell('INCOME STATEMENT', 'H27'))))

def calc_Ratio_Analysis_C13(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'H21')), 100), xl_ref(ctx.cell('INCOME STATEMENT', 'H6')))

def calc_Ratio_Analysis_C19(ctx):
    return xl_div(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'H21')), xl_ref(ctx.cell('INCOME STATEMENT', 'H27'))), xl_ref(ctx.cell('INCOME STATEMENT', 'H26')))

def calc_COMPANY_OVERVIEW_D22(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'O21')), xl_ref(ctx.cell('INCOME STATEMENT', 'O6')))

def calc_INCOME_STATEMENT_I22(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'I21')), xl_ref(ctx.cell('INCOME STATEMENT', 'O21'))), 100)

def calc_INCOME_STATEMENT_J22(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'J21')), xl_ref(ctx.cell('INCOME STATEMENT', 'O21'))), 100)

def calc_INCOME_STATEMENT_K22(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'K21')), xl_ref(ctx.cell('INCOME STATEMENT', 'O21'))), 100)

def calc_INCOME_STATEMENT_L22(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'L21')), xl_ref(ctx.cell('INCOME STATEMENT', 'O21'))), 100)

def calc_INCOME_STATEMENT_M22(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'M21')), xl_ref(ctx.cell('INCOME STATEMENT', 'O21'))), 100)

def calc_INCOME_STATEMENT_N22(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'N21')), xl_ref(ctx.cell('INCOME STATEMENT', 'O21'))), 100)

def calc_INCOME_STATEMENT_O23(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'O21')), xl_ref(ctx.cell('INCOME STATEMENT', 'H21'))), 1), 100)

def calc_INCOME_STATEMENT_O25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'O21')), xl_ref(ctx.cell('INCOME STATEMENT', 'O6'))), 100)

def calc_INCOME_STATEMENT_O28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'O21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'O26')), xl_ref(ctx.cell('INCOME STATEMENT', 'O27'))))

def calc_Ratio_Analysis_D13(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'O21')), 100), xl_ref(ctx.cell('INCOME STATEMENT', 'O6')))

def calc_Ratio_Analysis_D19(ctx):
    return xl_div(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'O21')), xl_ref(ctx.cell('INCOME STATEMENT', 'O27'))), xl_ref(ctx.cell('INCOME STATEMENT', 'O26')))

def calc_Ratio_Analysis_D22(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'O21')), xl_ref(ctx.cell('INCOME STATEMENT', 'H21'))), 1)

def calc_INCOME_STATEMENT_D32(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'D30')), xl_ref(ctx.cell('INCOME STATEMENT', 'D31')))

def calc_INCOME_STATEMENT_F30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'F28')), xl_ref(ctx.cell('INCOME STATEMENT', 'F29')))

def calc_INCOME_STATEMENT_M30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'M28')), xl_ref(ctx.cell('INCOME STATEMENT', 'M29')))

def calc_INCOME_STATEMENT_R32(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'R30')), xl_ref(ctx.cell('INCOME STATEMENT', 'R31')))

def calc_INCOME_STATEMENT_T30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'T28')), xl_ref(ctx.cell('INCOME STATEMENT', 'T29')))

def calc_COMPANY_OVERVIEW_E22(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'V21')), xl_ref(ctx.cell('INCOME STATEMENT', 'V6')))

def calc_COMPANY_OVERVIEW_E26(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'G13')), xl_ref(ctx.cell('INCOME STATEMENT', 'V21')))

def calc_INCOME_STATEMENT_P22(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'P21')), xl_ref(ctx.cell('INCOME STATEMENT', 'V21'))), 100)

def calc_INCOME_STATEMENT_Q22(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Q21')), xl_ref(ctx.cell('INCOME STATEMENT', 'V21'))), 100)

def calc_INCOME_STATEMENT_R22(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'R21')), xl_ref(ctx.cell('INCOME STATEMENT', 'V21'))), 100)

def calc_INCOME_STATEMENT_S22(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'S21')), xl_ref(ctx.cell('INCOME STATEMENT', 'V21'))), 100)

def calc_INCOME_STATEMENT_T22(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'T21')), xl_ref(ctx.cell('INCOME STATEMENT', 'V21'))), 100)

def calc_INCOME_STATEMENT_U22(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'U21')), xl_ref(ctx.cell('INCOME STATEMENT', 'V21'))), 100)

def calc_INCOME_STATEMENT_V23(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'V21')), xl_ref(ctx.cell('INCOME STATEMENT', 'O21'))), 1), 100)

def calc_INCOME_STATEMENT_V25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'V21')), xl_ref(ctx.cell('INCOME STATEMENT', 'V6'))), 100)

def calc_INCOME_STATEMENT_V28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'V21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'V26')), xl_ref(ctx.cell('INCOME STATEMENT', 'V27'))))

def calc_Valuation_B20(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'V21'))

def calc_Valuation_B21(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'V21')), xl_ref(ctx.cell('INCOME STATEMENT', 'V27')))

def calc_Ratio_Analysis_E13(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'V21')), 100), xl_ref(ctx.cell('INCOME STATEMENT', 'V6')))

def calc_Ratio_Analysis_E19(ctx):
    return xl_div(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'V21')), xl_ref(ctx.cell('INCOME STATEMENT', 'V27'))), xl_ref(ctx.cell('INCOME STATEMENT', 'V26')))

def calc_Ratio_Analysis_E22(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'V21')), xl_ref(ctx.cell('INCOME STATEMENT', 'O21'))), 1)

def calc_INCOME_STATEMENT_Y32(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'Y30')), xl_ref(ctx.cell('INCOME STATEMENT', 'Y31')))

def calc_INCOME_STATEMENT_AA30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AA28')), xl_ref(ctx.cell('INCOME STATEMENT', 'AA29')))

def calc_COMPANY_OVERVIEW_F22(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AC21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC6')))

def calc_COMPANY_OVERVIEW_F26(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'G13')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC21')))

def calc_INCOME_STATEMENT_W22(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'W21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC21'))), 100)

def calc_INCOME_STATEMENT_X22(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'X21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC21'))), 100)

def calc_INCOME_STATEMENT_Y22(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Y21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC21'))), 100)

def calc_INCOME_STATEMENT_Z22(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Z21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC21'))), 100)

def calc_INCOME_STATEMENT_AA22(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AA21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC21'))), 100)

def calc_INCOME_STATEMENT_AB22(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AB21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC21'))), 100)

def calc_INCOME_STATEMENT_AC23(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AC21')), xl_ref(ctx.cell('INCOME STATEMENT', 'V21'))), 1), 100)

def calc_INCOME_STATEMENT_AD23(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AD21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC21'))), 1), 100)

def calc_INCOME_STATEMENT_AC25(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AC21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC6'))), 100)

def calc_INCOME_STATEMENT_AC28(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AC21')), xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AC26')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC27'))))

def calc_Valuation_C20(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AC21'))

def calc_Ratio_Analysis_F13(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AC21')), 100), xl_ref(ctx.cell('INCOME STATEMENT', 'AC6')))

def calc_Ratio_Analysis_F19(ctx):
    return xl_div(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AC21')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC27'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AC26')))

def calc_Ratio_Analysis_F22(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AC21')), xl_ref(ctx.cell('INCOME STATEMENT', 'V21'))), 1)

def calc_INCOME_STATEMENT_AD32(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AD30')), xl_ref(ctx.cell('INCOME STATEMENT', 'AD31')))

def calc_Ratio_Analysis_G26(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AD30')), xl_ref(ctx.cell('INCOME STATEMENT', 'W30'))), 1)

def calc_Valuation_D44(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'D21')), xl_ref(ctx.cell('Valuation', 'D19')))

def calc_INCOME_STATEMENT_AE32(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AE30')), xl_ref(ctx.cell('INCOME STATEMENT', 'AE31')))

def calc_Ratio_Analysis_H26(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AE30')), xl_ref(ctx.cell('INCOME STATEMENT', 'X30'))), 1)

def calc_Valuation_E39(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'E21')), xl_ref(ctx.cell('Valuation', 'D21'))), 1)

def calc_Valuation_E44(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'E21')), xl_ref(ctx.cell('Valuation', 'E19')))

def calc_INCOME_STATEMENT_AF32(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AF30')), xl_ref(ctx.cell('INCOME STATEMENT', 'AF31')))

def calc_Ratio_Analysis_I26(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AF30')), xl_ref(ctx.cell('INCOME STATEMENT', 'Y30'))), 1)

def calc_Valuation_F39(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'F21')), xl_ref(ctx.cell('Valuation', 'E21'))), 1)

def calc_INCOME_STATEMENT_AG32(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AG30')), xl_ref(ctx.cell('INCOME STATEMENT', 'AG31')))

def calc_Ratio_Analysis_J26(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AG30')), xl_ref(ctx.cell('INCOME STATEMENT', 'Z30'))), 1)

def calc_Valuation_G39(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'G21')), xl_ref(ctx.cell('Valuation', 'F21'))), 1)

def calc_INCOME_STATEMENT_K41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'K38')), xl_ref(ctx.cell('INCOME STATEMENT', 'K39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'K40')))

def calc_INCOME_STATEMENT_H47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'G47'))

def calc_INCOME_STATEMENT_G48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'G47')), 10)

def calc_INCOME_STATEMENT_Y47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'X47'))

def calc_INCOME_STATEMENT_X48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'X47')), 10)

def calc_CASH_FOW_STATEMENT_I24(ctx):
    return xl_add(xl_ref(ctx.cell('CASH FOW STATEMENT', 'I22')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'I23')))

def calc_Valuation_H46(ctx):
    return xl_ref(ctx.cell('Valuation', 'G46'))

def calc_Valuation_K25(ctx):
    return xl_ref(ctx.cell('Valuation', 'J25'))

def calc_Valuation_H19(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'G19')), xl_add(1, xl_ref(ctx.cell('Valuation', 'H37'))))

def calc_Valuation_I37(ctx):
    return xl_sub(xl_ref(ctx.cell('Valuation', 'H37')), xl_div(xl_sub(xl_ref(ctx.cell('Valuation', 'E37')), xl_ref(ctx.cell('Valuation', 'J37'))), 5))

def calc_Valuation_I45(ctx):
    return xl_ref(ctx.cell('Valuation', 'H45'))

def calc_PRESENTATION_B45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'B44'))

def calc_PRESENTATION_C45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'C44'))

def calc_PRESENTATION_C47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'C37')), xl_ref(ctx.cell('PRESENTATION', 'C44')))

def calc_PRESENTATION_E45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'E44'))

def calc_PRESENTATION_E47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'E37')), xl_ref(ctx.cell('PRESENTATION', 'E44')))

def calc_PRESENTATION_G45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'G44'))

def calc_PRESENTATION_G47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'G37')), xl_ref(ctx.cell('PRESENTATION', 'G44')))

def calc_PRESENTATION_I45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'I44'))

def calc_PRESENTATION_I47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'I37')), xl_ref(ctx.cell('PRESENTATION', 'I44')))

def calc_PRESENTATION_J45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'J44'))

def calc_PRESENTATION_J47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'J37')), xl_ref(ctx.cell('PRESENTATION', 'J44')))

def calc_PRESENTATION_L45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'L44'))

def calc_PRESENTATION_L47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'L37')), xl_ref(ctx.cell('PRESENTATION', 'L44')))

def calc_PRESENTATION_N45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'N44'))

def calc_PRESENTATION_N47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'N37')), xl_ref(ctx.cell('PRESENTATION', 'N44')))

def calc_PRESENTATION_P45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'P44'))

def calc_PRESENTATION_P47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'P37')), xl_ref(ctx.cell('PRESENTATION', 'P44')))

def calc_PRESENTATION_Q45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'Q44'))

def calc_PRESENTATION_Q47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'Q37')), xl_ref(ctx.cell('PRESENTATION', 'Q44')))

def calc_PRESENTATION_S45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'S44'))

def calc_PRESENTATION_S47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'S37')), xl_ref(ctx.cell('PRESENTATION', 'S44')))

def calc_PRESENTATION_U45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'U44'))

def calc_PRESENTATION_U47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'U37')), xl_ref(ctx.cell('PRESENTATION', 'U44')))

def calc_PRESENTATION_W45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'W44'))

def calc_PRESENTATION_W47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'W37')), xl_ref(ctx.cell('PRESENTATION', 'W44')))

def calc_PRESENTATION_X45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'X44'))

def calc_PRESENTATION_X47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'X37')), xl_ref(ctx.cell('PRESENTATION', 'X44')))

def calc_PRESENTATION_Z45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'Z44'))

def calc_PRESENTATION_Z47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'Z37')), xl_ref(ctx.cell('PRESENTATION', 'Z44')))

def calc_PRESENTATION_AB45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'AB44'))

def calc_PRESENTATION_AB47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'AB37')), xl_ref(ctx.cell('PRESENTATION', 'AB44')))

def calc_PRESENTATION_AD44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AD40')), xl_ref(ctx.cell('PRESENTATION', 'AD41'))), xl_ref(ctx.cell('PRESENTATION', 'AD42'))), xl_ref(ctx.cell('PRESENTATION', 'AD43')))

def calc_PRESENTATION_AE44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AE40')), xl_ref(ctx.cell('PRESENTATION', 'AE41'))), xl_ref(ctx.cell('PRESENTATION', 'AE42'))), xl_ref(ctx.cell('PRESENTATION', 'AE43')))

def calc_PRESENTATION_AF44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AF40')), xl_ref(ctx.cell('PRESENTATION', 'AF41'))), xl_ref(ctx.cell('PRESENTATION', 'AF42'))), xl_ref(ctx.cell('PRESENTATION', 'AF43')))

def calc_PRESENTATION_AG44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AG40')), xl_ref(ctx.cell('PRESENTATION', 'AG41'))), xl_ref(ctx.cell('PRESENTATION', 'AG42'))), xl_ref(ctx.cell('PRESENTATION', 'AG43')))

def calc_PRESENTATION_H38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'H36')), xl_ref(ctx.cell('PRESENTATION', 'H37')))

def calc_PRESENTATION_O38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'O36')), xl_ref(ctx.cell('PRESENTATION', 'O37')))

def calc_PRESENTATION_D44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'D40')), xl_ref(ctx.cell('PRESENTATION', 'D41'))), xl_ref(ctx.cell('PRESENTATION', 'D42'))), xl_ref(ctx.cell('PRESENTATION', 'D43')))

def calc_PRESENTATION_F40(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'F38')), xl_ref(ctx.cell('PRESENTATION', 'F39')))

def calc_PRESENTATION_M40(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'M38')), xl_ref(ctx.cell('PRESENTATION', 'M39')))

def calc_PRESENTATION_R44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'R40')), xl_ref(ctx.cell('PRESENTATION', 'R41'))), xl_ref(ctx.cell('PRESENTATION', 'R42'))), xl_ref(ctx.cell('PRESENTATION', 'R43')))

def calc_PRESENTATION_T40(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'T38')), xl_ref(ctx.cell('PRESENTATION', 'T39')))

def calc_PRESENTATION_V38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'V36')), xl_ref(ctx.cell('PRESENTATION', 'V37')))

def calc_PRESENTATION_Y44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'Y40')), xl_ref(ctx.cell('PRESENTATION', 'Y41'))), xl_ref(ctx.cell('PRESENTATION', 'Y42'))), xl_ref(ctx.cell('PRESENTATION', 'Y43')))

def calc_PRESENTATION_AA40(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AA38')), xl_ref(ctx.cell('PRESENTATION', 'AA39')))

def calc_PRESENTATION_AC38(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'AC36')), xl_ref(ctx.cell('PRESENTATION', 'AC37')))

def calc_Segment_Revenue_Model_U33(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'U17')), xl_ref(ctx.cell('Segment Revenue Model', 'P17'))), 1)

def calc_Segment_Revenue_Model_V33(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'V17')), xl_ref(ctx.cell('Segment Revenue Model', 'U17'))), 1)

def calc_Segment_Revenue_Model_U37(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'U8')), xl_ref(ctx.cell('Segment Revenue Model', 'U17')))

def calc_Segment_Revenue_Model_U38(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'U9')), xl_ref(ctx.cell('Segment Revenue Model', 'U17')))

def calc_Segment_Revenue_Model_U39(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'U10')), xl_ref(ctx.cell('Segment Revenue Model', 'U17')))

def calc_Segment_Revenue_Model_U40(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'U11')), xl_ref(ctx.cell('Segment Revenue Model', 'U17')))

def calc_Segment_Revenue_Model_U42(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'U13')), xl_ref(ctx.cell('Segment Revenue Model', 'U17')))

def calc_Segment_Revenue_Model_U43(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'U14')), xl_ref(ctx.cell('Segment Revenue Model', 'U17')))

def calc_Segment_Revenue_Model_U44(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'U15')), xl_ref(ctx.cell('Segment Revenue Model', 'U17')))

def calc_Segment_Revenue_Model_U45(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'U16')), xl_ref(ctx.cell('Segment Revenue Model', 'U17')))

def calc_Segment_Revenue_Model_U46(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'U17')), xl_ref(ctx.cell('Segment Revenue Model', 'U17')))

def calc_Segment_Revenue_Model_Y15(ctx):
    return xl_mul(xl_add(xl_ref(ctx.cell('Segment Revenue Model', 'Y31')), 1), xl_ref(ctx.cell('Segment Revenue Model', 'X15')))

def calc_Segment_Revenue_Model_X17(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'X15')), xl_ref(ctx.cell('Segment Revenue Model', 'X16')))

def calc_Segment_Revenue_Model_W33(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'W17')), xl_ref(ctx.cell('Segment Revenue Model', 'V17'))), 1)

def calc_Segment_Revenue_Model_W37(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'W8')), xl_ref(ctx.cell('Segment Revenue Model', 'W17')))

def calc_Segment_Revenue_Model_W38(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'W9')), xl_ref(ctx.cell('Segment Revenue Model', 'W17')))

def calc_Segment_Revenue_Model_W39(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'W10')), xl_ref(ctx.cell('Segment Revenue Model', 'W17')))

def calc_Segment_Revenue_Model_W40(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'W11')), xl_ref(ctx.cell('Segment Revenue Model', 'W17')))

def calc_Segment_Revenue_Model_W41(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'W12')), xl_ref(ctx.cell('Segment Revenue Model', 'W17')))

def calc_Segment_Revenue_Model_W42(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'W13')), xl_ref(ctx.cell('Segment Revenue Model', 'W17')))

def calc_Segment_Revenue_Model_W43(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'W14')), xl_ref(ctx.cell('Segment Revenue Model', 'W17')))

def calc_Segment_Revenue_Model_W44(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'W15')), xl_ref(ctx.cell('Segment Revenue Model', 'W17')))

def calc_Segment_Revenue_Model_W45(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'W16')), xl_ref(ctx.cell('Segment Revenue Model', 'W17')))

def calc_Segment_Revenue_Model_W46(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'W17')), xl_ref(ctx.cell('Segment Revenue Model', 'W17')))

def calc_INCOME_STATEMENT_B41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'B38')), xl_ref(ctx.cell('INCOME STATEMENT', 'B39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'B40')))

def calc_INCOME_STATEMENT_B43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'B38')), xl_error('#REF!')), 1), 100)

def calc_INCOME_STATEMENT_B44(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'B38')), xl_error('#REF!')), 1), 100)

def calc_INCOME_STATEMENT_C41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'C38')), xl_ref(ctx.cell('INCOME STATEMENT', 'C39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'C40')))

def calc_INCOME_STATEMENT_C43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'C38')), xl_error('#REF!')), 1), 100)

def calc_INCOME_STATEMENT_C44(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'C38')), xl_ref(ctx.cell('INCOME STATEMENT', 'B38'))), 1), 100)

def calc_INCOME_STATEMENT_E41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'E38')), xl_ref(ctx.cell('INCOME STATEMENT', 'E39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'E40')))

def calc_INCOME_STATEMENT_E43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'E38')), xl_error('#REF!')), 1), 100)

def calc_INCOME_STATEMENT_E44(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'E38')), xl_ref(ctx.cell('INCOME STATEMENT', 'C38'))), 1), 100)

def calc_INCOME_STATEMENT_G41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'G38')), xl_ref(ctx.cell('INCOME STATEMENT', 'G39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'G40')))

def calc_INCOME_STATEMENT_G43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'G38')), xl_error('#REF!')), 1), 100)

def calc_INCOME_STATEMENT_G44(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'G38')), xl_ref(ctx.cell('INCOME STATEMENT', 'E38'))), 1), 100)

def calc_INCOME_STATEMENT_I41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'I38')), xl_ref(ctx.cell('INCOME STATEMENT', 'I39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'I40')))

def calc_INCOME_STATEMENT_I43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'I38')), xl_ref(ctx.cell('INCOME STATEMENT', 'B38'))), 1), 100)

def calc_INCOME_STATEMENT_I44(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'I38')), xl_ref(ctx.cell('INCOME STATEMENT', 'G38'))), 1), 100)

def calc_INCOME_STATEMENT_J41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'J38')), xl_ref(ctx.cell('INCOME STATEMENT', 'J39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'J40')))

def calc_INCOME_STATEMENT_J43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'J38')), xl_ref(ctx.cell('INCOME STATEMENT', 'C38'))), 1), 100)

def calc_INCOME_STATEMENT_J44(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'J38')), xl_ref(ctx.cell('INCOME STATEMENT', 'I38'))), 1), 100)

def calc_INCOME_STATEMENT_L41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'L38')), xl_ref(ctx.cell('INCOME STATEMENT', 'L39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'L40')))

def calc_INCOME_STATEMENT_L43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'L38')), xl_ref(ctx.cell('INCOME STATEMENT', 'E38'))), 1), 100)

def calc_INCOME_STATEMENT_L44(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'L38')), xl_ref(ctx.cell('INCOME STATEMENT', 'J38'))), 1), 100)

def calc_INCOME_STATEMENT_N41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'N38')), xl_ref(ctx.cell('INCOME STATEMENT', 'N39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'N40')))

def calc_INCOME_STATEMENT_N43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'N38')), xl_ref(ctx.cell('INCOME STATEMENT', 'G38'))), 1), 100)

def calc_INCOME_STATEMENT_N44(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'N38')), xl_ref(ctx.cell('INCOME STATEMENT', 'L38'))), 1), 100)

def calc_INCOME_STATEMENT_P41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'P38')), xl_ref(ctx.cell('INCOME STATEMENT', 'P39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'P40')))

def calc_INCOME_STATEMENT_P43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'P38')), xl_ref(ctx.cell('INCOME STATEMENT', 'I38'))), 1), 100)

def calc_INCOME_STATEMENT_P44(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'P38')), xl_ref(ctx.cell('INCOME STATEMENT', 'N38'))), 1), 100)

def calc_INCOME_STATEMENT_Q41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'Q38')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'Q40')))

def calc_INCOME_STATEMENT_Q43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Q38')), xl_ref(ctx.cell('INCOME STATEMENT', 'J38'))), 1), 100)

def calc_INCOME_STATEMENT_Q44(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Q38')), xl_ref(ctx.cell('INCOME STATEMENT', 'P38'))), 1), 100)

def calc_INCOME_STATEMENT_S41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'S38')), xl_ref(ctx.cell('INCOME STATEMENT', 'S39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'S40')))

def calc_INCOME_STATEMENT_S43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'S38')), xl_ref(ctx.cell('INCOME STATEMENT', 'L38'))), 1), 100)

def calc_INCOME_STATEMENT_S44(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'S38')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q38'))), 1), 100)

def calc_INCOME_STATEMENT_U41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'U38')), xl_ref(ctx.cell('INCOME STATEMENT', 'U39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'U40')))

def calc_INCOME_STATEMENT_U43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'U38')), xl_ref(ctx.cell('INCOME STATEMENT', 'N38'))), 1), 100)

def calc_INCOME_STATEMENT_U44(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'U38')), xl_ref(ctx.cell('INCOME STATEMENT', 'S38'))), 1), 100)

def calc_INCOME_STATEMENT_W41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'W38')), xl_ref(ctx.cell('INCOME STATEMENT', 'W39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'W40')))

def calc_INCOME_STATEMENT_W43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'W38')), xl_ref(ctx.cell('INCOME STATEMENT', 'P38'))), 1), 100)

def calc_INCOME_STATEMENT_W44(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'W38')), xl_ref(ctx.cell('INCOME STATEMENT', 'U38'))), 1), 100)

def calc_INCOME_STATEMENT_X41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'X38')), xl_ref(ctx.cell('INCOME STATEMENT', 'X39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'X40')))

def calc_INCOME_STATEMENT_X43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'X38')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q38'))), 1), 100)

def calc_INCOME_STATEMENT_X44(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'X38')), xl_ref(ctx.cell('INCOME STATEMENT', 'W38'))), 1), 100)

def calc_INCOME_STATEMENT_Z41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'Z38')), xl_ref(ctx.cell('INCOME STATEMENT', 'Z39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'Z40')))

def calc_INCOME_STATEMENT_Z43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Z38')), xl_ref(ctx.cell('INCOME STATEMENT', 'S38'))), 1), 100)

def calc_INCOME_STATEMENT_Z44(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Z38')), xl_ref(ctx.cell('INCOME STATEMENT', 'X38'))), 1), 100)

def calc_INCOME_STATEMENT_AB41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AB38')), xl_ref(ctx.cell('INCOME STATEMENT', 'AB39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AB40')))

def calc_INCOME_STATEMENT_AB43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AB38')), xl_ref(ctx.cell('INCOME STATEMENT', 'U38'))), 1), 100)

def calc_INCOME_STATEMENT_AB44(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AB38')), xl_ref(ctx.cell('INCOME STATEMENT', 'Z38'))), 1), 100)

def calc_INCOME_STATEMENT_H22(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'F22')), xl_ref(ctx.cell('INCOME STATEMENT', 'G22')))

def calc_INCOME_STATEMENT_H30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'H28')), xl_ref(ctx.cell('INCOME STATEMENT', 'H29')))

def calc_PRESENTATION_C63(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'C13'))

def calc_COMPANY_OVERVIEW_C24(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'C19'))

def calc_PRESENTATION_C66(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'C19'))

def calc_INCOME_STATEMENT_O22(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'M22')), xl_ref(ctx.cell('INCOME STATEMENT', 'N22')))

def calc_INCOME_STATEMENT_O30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'O28')), xl_ref(ctx.cell('INCOME STATEMENT', 'O29')))

def calc_PRESENTATION_D63(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'D13'))

def calc_COMPANY_OVERVIEW_D24(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'D19'))

def calc_PRESENTATION_D66(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'D19'))

def calc_INCOME_STATEMENT_D36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'D33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'D32'))))

def calc_INCOME_STATEMENT_D37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'D33')), xl_ref(ctx.cell('INCOME STATEMENT', 'D34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'D35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'D32'))), 100)

def calc_INCOME_STATEMENT_D38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'D32')), xl_ref(ctx.cell('INCOME STATEMENT', 'D33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'D34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'D35')))

def calc_INCOME_STATEMENT_F32(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'F30')), xl_ref(ctx.cell('INCOME STATEMENT', 'F31')))

def calc_INCOME_STATEMENT_M32(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'M30')), xl_ref(ctx.cell('INCOME STATEMENT', 'M31')))

def calc_INCOME_STATEMENT_R36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'R33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'R32'))))

def calc_INCOME_STATEMENT_R37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'R33')), xl_ref(ctx.cell('INCOME STATEMENT', 'R34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'R35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'R32'))), 100)

def calc_INCOME_STATEMENT_R38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'R32')), xl_ref(ctx.cell('INCOME STATEMENT', 'R33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'R34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'R35')))

def calc_INCOME_STATEMENT_T32(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'T30')), xl_ref(ctx.cell('INCOME STATEMENT', 'T31')))

def calc_INCOME_STATEMENT_V22(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'T22')), xl_ref(ctx.cell('INCOME STATEMENT', 'U22')))

def calc_INCOME_STATEMENT_V30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'V28')), xl_ref(ctx.cell('INCOME STATEMENT', 'V29')))

def calc_Valuation_B43(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'B20')), xl_ref(ctx.cell('Valuation', 'B19')))

def calc_Valuation_B23(ctx):
    return xl_sub(xl_ref(ctx.cell('Valuation', 'B21')), xl_ref(ctx.cell('Valuation', 'B22')))

def calc_Valuation_B44(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'B21')), xl_ref(ctx.cell('Valuation', 'B19')))

def calc_Valuation_B48(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'B22')), xl_ref(ctx.cell('Valuation', 'B21')))

def calc_PRESENTATION_E63(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'E13'))

def calc_COMPANY_OVERVIEW_E24(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'E19'))

def calc_PRESENTATION_E66(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'E19'))

def calc_INCOME_STATEMENT_Y36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Y33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'Y32'))))

def calc_INCOME_STATEMENT_Y37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'Y33')), xl_ref(ctx.cell('INCOME STATEMENT', 'Y34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'Y35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'Y32'))), 100)

def calc_INCOME_STATEMENT_Y38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'Y32')), xl_ref(ctx.cell('INCOME STATEMENT', 'Y33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'Y34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'Y35')))

def calc_INCOME_STATEMENT_AA32(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AA30')), xl_ref(ctx.cell('INCOME STATEMENT', 'AA31')))

def calc_INCOME_STATEMENT_AC22(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AA22')), xl_ref(ctx.cell('INCOME STATEMENT', 'AB22')))

def calc_INCOME_STATEMENT_AC30(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AC28')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC29')))

def calc_Valuation_C21(ctx):
    return xl_sub(xl_ref(ctx.cell('Valuation', 'C20')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC27')))

def calc_Valuation_C38(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'C20')), xl_ref(ctx.cell('Valuation', 'B20'))), 1)

def calc_Valuation_D38(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'D20')), xl_ref(ctx.cell('Valuation', 'C20'))), 1)

def calc_Valuation_C43(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'C20')), xl_ref(ctx.cell('Valuation', 'C19')))

def calc_PRESENTATION_F63(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'F13'))

def calc_COMPANY_OVERVIEW_F24(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'F19'))

def calc_PRESENTATION_F66(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'F19'))

def calc_INCOME_STATEMENT_AD33(ctx):
    return xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AD32')), xl_percent(33))

def calc_Ratio_Analysis_G12(ctx):
    return xl_div(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AD32')), xl_ref(ctx.cell('INCOME STATEMENT', 'AD26'))), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'G8')), xl_ref(ctx.cell('BALANCESHEET', 'F8'))))

def calc_Ratio_Analysis_G18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AD29')), xl_ref(ctx.cell('INCOME STATEMENT', 'AD32')))

def calc_INCOME_STATEMENT_AE33(ctx):
    return xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AE32')), xl_percent(33))

def calc_CASH_FOW_STATEMENT_F8(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AE32'))

def calc_Ratio_Analysis_H12(ctx):
    return xl_div(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AE32')), xl_ref(ctx.cell('INCOME STATEMENT', 'AE26'))), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'H8')), xl_ref(ctx.cell('BALANCESHEET', 'G8'))))

def calc_Ratio_Analysis_H18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AE29')), xl_ref(ctx.cell('INCOME STATEMENT', 'AE32')))

def calc_Valuation_F44(ctx):
    return xl_ref(ctx.cell('Valuation', 'E44'))

def calc_INCOME_STATEMENT_AF33(ctx):
    return xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AF32')), xl_percent(33))

def calc_CASH_FOW_STATEMENT_G8(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AF32'))

def calc_Ratio_Analysis_I12(ctx):
    return xl_div(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AF32')), xl_ref(ctx.cell('INCOME STATEMENT', 'AF26'))), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'I8')), xl_ref(ctx.cell('BALANCESHEET', 'H8'))))

def calc_Ratio_Analysis_I18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AF29')), xl_ref(ctx.cell('INCOME STATEMENT', 'AF32')))

def calc_INCOME_STATEMENT_AG33(ctx):
    return xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AG32')), xl_percent(33))

def calc_CASH_FOW_STATEMENT_H8(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AG32'))

def calc_Ratio_Analysis_J12(ctx):
    return xl_div(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AG32')), xl_ref(ctx.cell('INCOME STATEMENT', 'AG26'))), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'J8')), xl_ref(ctx.cell('BALANCESHEET', 'I8'))))

def calc_Ratio_Analysis_J18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AG29')), xl_ref(ctx.cell('INCOME STATEMENT', 'AG32')))

def calc_INCOME_STATEMENT_K45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'K41')), xl_ref(ctx.cell('INCOME STATEMENT', 'L6'))), 100)

def calc_INCOME_STATEMENT_I47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'H47'))

def calc_INCOME_STATEMENT_H48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'H47')), 10)

def calc_INCOME_STATEMENT_Z47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'Y47'))

def calc_INCOME_STATEMENT_Y48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Y47')), 10)

def calc_Valuation_I46(ctx):
    return xl_ref(ctx.cell('Valuation', 'H46'))

def calc_Valuation_H24(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'H45')), xl_ref(ctx.cell('Valuation', 'H19')))

def calc_Valuation_H27(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'H46')), xl_uminus(xl_ref(ctx.cell('Valuation', 'H19'))))

def calc_Valuation_I19(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'H19')), xl_add(1, xl_ref(ctx.cell('Valuation', 'I37'))))

def calc_Valuation_J45(ctx):
    return xl_ref(ctx.cell('Valuation', 'I45'))

def calc_PRESENTATION_B47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'B46')), xl_ref(ctx.cell('PRESENTATION', 'B45')))

def calc_PRESENTATION_AD45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'AD44'))

def calc_PRESENTATION_AD47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'AD37')), xl_ref(ctx.cell('PRESENTATION', 'AD44')))

def calc_PRESENTATION_AE45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'AE44'))

def calc_PRESENTATION_AE47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'AE37')), xl_ref(ctx.cell('PRESENTATION', 'AE44')))

def calc_PRESENTATION_AF45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'AF44'))

def calc_PRESENTATION_AF47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'AF37')), xl_ref(ctx.cell('PRESENTATION', 'AF44')))

def calc_PRESENTATION_AG45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'AG44'))

def calc_PRESENTATION_AG47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'AG37')), xl_ref(ctx.cell('PRESENTATION', 'AG44')))

def calc_PRESENTATION_H40(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'H38')), xl_ref(ctx.cell('PRESENTATION', 'H39')))

def calc_PRESENTATION_O40(ctx):
    return xl_add(xl_ref(ctx.cell('PRESENTATION', 'O38')), xl_ref(ctx.cell('PRESENTATION', 'O39')))

def calc_PRESENTATION_D45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'D44'))

def calc_PRESENTATION_D47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'D37')), xl_ref(ctx.cell('PRESENTATION', 'D44')))

def calc_PRESENTATION_F44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'F40')), xl_ref(ctx.cell('PRESENTATION', 'F41'))), xl_ref(ctx.cell('PRESENTATION', 'F42'))), xl_ref(ctx.cell('PRESENTATION', 'F43')))

def calc_PRESENTATION_M44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'M40')), xl_ref(ctx.cell('PRESENTATION', 'M41'))), xl_ref(ctx.cell('PRESENTATION', 'M42'))), xl_ref(ctx.cell('PRESENTATION', 'M43')))

def calc_PRESENTATION_R45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'R44'))

def calc_PRESENTATION_R47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'R37')), xl_ref(ctx.cell('PRESENTATION', 'R44')))

def calc_PRESENTATION_T44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'T40')), xl_ref(ctx.cell('PRESENTATION', 'T41'))), xl_ref(ctx.cell('PRESENTATION', 'T42'))), xl_ref(ctx.cell('PRESENTATION', 'T43')))

def calc_PRESENTATION_V40(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'V38')), xl_ref(ctx.cell('PRESENTATION', 'V39')))

def calc_PRESENTATION_Y45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'Y44'))

def calc_PRESENTATION_Y47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'Y37')), xl_ref(ctx.cell('PRESENTATION', 'Y44')))

def calc_PRESENTATION_AA44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AA40')), xl_ref(ctx.cell('PRESENTATION', 'AA41'))), xl_ref(ctx.cell('PRESENTATION', 'AA42'))), xl_ref(ctx.cell('PRESENTATION', 'AA43')))

def calc_PRESENTATION_AC40(ctx):
    return xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AC38')), xl_ref(ctx.cell('PRESENTATION', 'AC39')))

def calc_Segment_Revenue_Model_Y17(ctx):
    return xl_sub(xl_ref(ctx.cell('Segment Revenue Model', 'Y15')), xl_ref(ctx.cell('Segment Revenue Model', 'Y16')))

def calc_Segment_Revenue_Model_X33(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'X17')), xl_ref(ctx.cell('Segment Revenue Model', 'W17'))), 1)

def calc_Segment_Revenue_Model_X37(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'X8')), xl_ref(ctx.cell('Segment Revenue Model', 'X17')))

def calc_Segment_Revenue_Model_X38(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'X9')), xl_ref(ctx.cell('Segment Revenue Model', 'X17')))

def calc_Segment_Revenue_Model_X39(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'X10')), xl_ref(ctx.cell('Segment Revenue Model', 'X17')))

def calc_Segment_Revenue_Model_X40(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'X11')), xl_ref(ctx.cell('Segment Revenue Model', 'X17')))

def calc_Segment_Revenue_Model_X41(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'X12')), xl_ref(ctx.cell('Segment Revenue Model', 'X17')))

def calc_Segment_Revenue_Model_X42(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'X13')), xl_ref(ctx.cell('Segment Revenue Model', 'X17')))

def calc_Segment_Revenue_Model_X43(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'X14')), xl_ref(ctx.cell('Segment Revenue Model', 'X17')))

def calc_Segment_Revenue_Model_X44(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'X15')), xl_ref(ctx.cell('Segment Revenue Model', 'X17')))

def calc_Segment_Revenue_Model_X45(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'X16')), xl_ref(ctx.cell('Segment Revenue Model', 'X17')))

def calc_Segment_Revenue_Model_X46(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'X17')), xl_ref(ctx.cell('Segment Revenue Model', 'X17')))

def calc_INCOME_STATEMENT_B45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'B41')), xl_ref(ctx.cell('INCOME STATEMENT', 'B6'))), 100)

def calc_INCOME_STATEMENT_B49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'B41')), xl_ref(ctx.cell('INCOME STATEMENT', 'B48')))

def calc_INCOME_STATEMENT_C45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'C41')), xl_ref(ctx.cell('INCOME STATEMENT', 'C6'))), 100)

def calc_INCOME_STATEMENT_C49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'C41')), xl_ref(ctx.cell('INCOME STATEMENT', 'C48')))

def calc_INCOME_STATEMENT_E45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'E41')), xl_ref(ctx.cell('INCOME STATEMENT', 'E6'))), 100)

def calc_INCOME_STATEMENT_E49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'E41')), xl_ref(ctx.cell('INCOME STATEMENT', 'E48')))

def calc_INCOME_STATEMENT_G45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'G41')), xl_ref(ctx.cell('INCOME STATEMENT', 'G6'))), 100)

def calc_INCOME_STATEMENT_G49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'G41')), xl_ref(ctx.cell('INCOME STATEMENT', 'G48')))

def calc_INCOME_STATEMENT_I45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'I41')), xl_ref(ctx.cell('INCOME STATEMENT', 'I6'))), 100)

def calc_INCOME_STATEMENT_J45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'J41')), xl_ref(ctx.cell('INCOME STATEMENT', 'J6'))), 100)

def calc_INCOME_STATEMENT_L45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'L41')), xl_ref(ctx.cell('INCOME STATEMENT', 'L6'))), 100)

def calc_INCOME_STATEMENT_N45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'N41')), xl_ref(ctx.cell('INCOME STATEMENT', 'N6'))), 100)

def calc_INCOME_STATEMENT_P45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'P41')), xl_ref(ctx.cell('INCOME STATEMENT', 'P6'))), 100)

def calc_INCOME_STATEMENT_Q45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Q41')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q6'))), 100)

def calc_INCOME_STATEMENT_S45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'S41')), xl_ref(ctx.cell('INCOME STATEMENT', 'S6'))), 100)

def calc_INCOME_STATEMENT_S49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'S41')), xl_ref(ctx.cell('INCOME STATEMENT', 'S48')))

def calc_INCOME_STATEMENT_U45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'U41')), xl_ref(ctx.cell('INCOME STATEMENT', 'U6'))), 100)

def calc_INCOME_STATEMENT_U49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'U41')), xl_ref(ctx.cell('INCOME STATEMENT', 'U48')))

def calc_INCOME_STATEMENT_W45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'W41')), xl_ref(ctx.cell('INCOME STATEMENT', 'W6'))), 100)

def calc_INCOME_STATEMENT_W49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'W41')), xl_ref(ctx.cell('INCOME STATEMENT', 'W48')))

def calc_INCOME_STATEMENT_X45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'X41')), xl_ref(ctx.cell('INCOME STATEMENT', 'X6'))), 100)

def calc_INCOME_STATEMENT_X49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'X41')), xl_ref(ctx.cell('INCOME STATEMENT', 'X48')))

def calc_INCOME_STATEMENT_Z45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Z41')), xl_ref(ctx.cell('INCOME STATEMENT', 'Z6'))), 100)

def calc_INCOME_STATEMENT_AB45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AB41')), xl_ref(ctx.cell('INCOME STATEMENT', 'AB6'))), 100)

def calc_INCOME_STATEMENT_H32(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'H30')), xl_ref(ctx.cell('INCOME STATEMENT', 'H31')))

def calc_INCOME_STATEMENT_O32(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'O30')), xl_ref(ctx.cell('INCOME STATEMENT', 'O31')))

def calc_Ratio_Analysis_D26(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'O30')), xl_ref(ctx.cell('INCOME STATEMENT', 'H30'))), 1)

def calc_INCOME_STATEMENT_D41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'D38')), xl_ref(ctx.cell('INCOME STATEMENT', 'D39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'D40')))

def calc_INCOME_STATEMENT_D43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'D38')), xl_error('#REF!')), 1), 100)

def calc_INCOME_STATEMENT_K43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'K38')), xl_ref(ctx.cell('INCOME STATEMENT', 'D38'))), 1), 100)

def calc_INCOME_STATEMENT_F36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'F33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'F32'))))

def calc_INCOME_STATEMENT_F37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'F33')), xl_ref(ctx.cell('INCOME STATEMENT', 'F34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'F35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'F32'))), 100)

def calc_INCOME_STATEMENT_F38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'F32')), xl_ref(ctx.cell('INCOME STATEMENT', 'F33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'F34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'F35')))

def calc_INCOME_STATEMENT_M36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'M33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'M32'))))

def calc_INCOME_STATEMENT_M37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'M33')), xl_ref(ctx.cell('INCOME STATEMENT', 'M34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'M35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'M32'))), 100)

def calc_INCOME_STATEMENT_M38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'M32')), xl_ref(ctx.cell('INCOME STATEMENT', 'M33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'M34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'M35')))

def calc_INCOME_STATEMENT_R41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'R38')), xl_ref(ctx.cell('INCOME STATEMENT', 'R39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'R40')))

def calc_INCOME_STATEMENT_R43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'R38')), xl_ref(ctx.cell('INCOME STATEMENT', 'K38'))), 1), 100)

def calc_INCOME_STATEMENT_T36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'T33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'T32'))))

def calc_INCOME_STATEMENT_T37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'T33')), xl_ref(ctx.cell('INCOME STATEMENT', 'T34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'T35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'T32'))), 100)

def calc_INCOME_STATEMENT_T38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'T32')), xl_ref(ctx.cell('INCOME STATEMENT', 'T33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'T34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'T35')))

def calc_INCOME_STATEMENT_V32(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'V30')), xl_ref(ctx.cell('INCOME STATEMENT', 'V31')))

def calc_Ratio_Analysis_E26(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'V30')), xl_ref(ctx.cell('INCOME STATEMENT', 'O30'))), 1)

def calc_Valuation_B26(ctx):
    return xl_sub(xl_add(xl_ref(ctx.cell('Valuation', 'B23')), xl_ref(ctx.cell('Valuation', 'B24'))), xl_ref(ctx.cell('Valuation', 'B25')))

def calc_INCOME_STATEMENT_Y41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'Y38')), xl_ref(ctx.cell('INCOME STATEMENT', 'Y39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'Y40')))

def calc_INCOME_STATEMENT_Y43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Y38')), xl_ref(ctx.cell('INCOME STATEMENT', 'R38'))), 1), 100)

def calc_INCOME_STATEMENT_AA36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AA33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AA32'))))

def calc_INCOME_STATEMENT_AA37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AA33')), xl_ref(ctx.cell('INCOME STATEMENT', 'AA34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AA35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AA32'))), 100)

def calc_INCOME_STATEMENT_AA38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AA32')), xl_ref(ctx.cell('INCOME STATEMENT', 'AA33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AA34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AA35')))

def calc_INCOME_STATEMENT_AC32(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AC30')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC31')))

def calc_Ratio_Analysis_F26(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AC30')), xl_ref(ctx.cell('INCOME STATEMENT', 'V30'))), 1)

def calc_Valuation_C23(ctx):
    return xl_sub(xl_ref(ctx.cell('Valuation', 'C21')), xl_ref(ctx.cell('Valuation', 'C22')))

def calc_Valuation_C39(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'C21')), xl_ref(ctx.cell('Valuation', 'B21'))), 1)

def calc_Valuation_D39(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'D21')), xl_ref(ctx.cell('Valuation', 'C21'))), 1)

def calc_Valuation_C44(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'C21')), xl_ref(ctx.cell('Valuation', 'C19')))

def calc_Valuation_C48(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'C22')), xl_ref(ctx.cell('Valuation', 'C21')))

def calc_INCOME_STATEMENT_AD36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AD33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AD32'))))

def calc_INCOME_STATEMENT_AD37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AD33')), xl_ref(ctx.cell('INCOME STATEMENT', 'AD34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AD35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AD32'))), 100)

def calc_INCOME_STATEMENT_AD38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AD32')), xl_ref(ctx.cell('INCOME STATEMENT', 'AD33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AD34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AD35')))

def calc_Valuation_D22(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AD33'))

def calc_INCOME_STATEMENT_AE36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AE33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AE32'))))

def calc_INCOME_STATEMENT_AE37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AE33')), xl_ref(ctx.cell('INCOME STATEMENT', 'AE34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AE35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AE32'))), 100)

def calc_INCOME_STATEMENT_AE38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AE32')), xl_ref(ctx.cell('INCOME STATEMENT', 'AE33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AE34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AE35')))

def calc_CASH_FOW_STATEMENT_F9(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AE33')), xl_ref(ctx.cell('INCOME STATEMENT', 'AE34')))

def calc_Valuation_E22(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AE33'))

def calc_Valuation_G44(ctx):
    return xl_ref(ctx.cell('Valuation', 'F44'))

def calc_INCOME_STATEMENT_AF36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AF33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AF32'))))

def calc_INCOME_STATEMENT_AF37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AF33')), xl_ref(ctx.cell('INCOME STATEMENT', 'AF34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AF35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AF32'))), 100)

def calc_INCOME_STATEMENT_AF38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AF32')), xl_ref(ctx.cell('INCOME STATEMENT', 'AF33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AF34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AF35')))

def calc_CASH_FOW_STATEMENT_G9(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AF33')), xl_ref(ctx.cell('INCOME STATEMENT', 'AF34')))

def calc_Valuation_F22(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AF33'))

def calc_INCOME_STATEMENT_AG36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AG33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AG32'))))

def calc_INCOME_STATEMENT_AG37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AG33')), xl_ref(ctx.cell('INCOME STATEMENT', 'AG34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AG35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AG32'))), 100)

def calc_INCOME_STATEMENT_AG38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AG32')), xl_ref(ctx.cell('INCOME STATEMENT', 'AG33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AG34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AG35')))

def calc_CASH_FOW_STATEMENT_H9(ctx):
    return xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AG33')), xl_ref(ctx.cell('INCOME STATEMENT', 'AG34')))

def calc_Valuation_G22(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AG33'))

def calc_INCOME_STATEMENT_J47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'I47'))

def calc_INCOME_STATEMENT_I48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'I47')), 10)

def calc_INCOME_STATEMENT_AA47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'Z47'))

def calc_INCOME_STATEMENT_Z48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Z47')), 10)

def calc_Valuation_J46(ctx):
    return xl_ref(ctx.cell('Valuation', 'I46'))

def calc_Valuation_H40(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'H24')), xl_ref(ctx.cell('Valuation', 'G24'))), 1)

def calc_Valuation_J19(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'I19')), xl_add(1, xl_ref(ctx.cell('Valuation', 'J37'))))

def calc_Valuation_I24(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'I45')), xl_ref(ctx.cell('Valuation', 'I19')))

def calc_Valuation_I27(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'I46')), xl_uminus(xl_ref(ctx.cell('Valuation', 'I19'))))

def calc_Valuation_K45(ctx):
    return xl_ref(ctx.cell('Valuation', 'J45'))

def calc_PRESENTATION_H44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'H40')), xl_ref(ctx.cell('PRESENTATION', 'H41'))), xl_ref(ctx.cell('PRESENTATION', 'H42'))), xl_ref(ctx.cell('PRESENTATION', 'H43')))

def calc_PRESENTATION_O44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'O40')), xl_ref(ctx.cell('PRESENTATION', 'O41'))), xl_ref(ctx.cell('PRESENTATION', 'O42'))), xl_ref(ctx.cell('PRESENTATION', 'O43')))

def calc_PRESENTATION_F45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'F44'))

def calc_PRESENTATION_F47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'F37')), xl_ref(ctx.cell('PRESENTATION', 'F44')))

def calc_PRESENTATION_M45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'M44'))

def calc_PRESENTATION_M47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'M37')), xl_ref(ctx.cell('PRESENTATION', 'M44')))

def calc_PRESENTATION_T45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'T44'))

def calc_PRESENTATION_T47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'T37')), xl_ref(ctx.cell('PRESENTATION', 'T44')))

def calc_PRESENTATION_V44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'V40')), xl_ref(ctx.cell('PRESENTATION', 'V41'))), xl_ref(ctx.cell('PRESENTATION', 'V42'))), xl_ref(ctx.cell('PRESENTATION', 'V43')))

def calc_PRESENTATION_AA45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'AA44'))

def calc_PRESENTATION_AA47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'AA37')), xl_ref(ctx.cell('PRESENTATION', 'AA44')))

def calc_PRESENTATION_AC44(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('PRESENTATION', 'AC40')), xl_ref(ctx.cell('PRESENTATION', 'AC41'))), xl_ref(ctx.cell('PRESENTATION', 'AC42'))), xl_ref(ctx.cell('PRESENTATION', 'AC43')))

def calc_Segment_Revenue_Model_Y33(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Y17')), xl_ref(ctx.cell('Segment Revenue Model', 'X17'))), 1)

def calc_Segment_Revenue_Model_Y37(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Y8')), xl_ref(ctx.cell('Segment Revenue Model', 'Y17')))

def calc_Segment_Revenue_Model_Y38(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Y9')), xl_ref(ctx.cell('Segment Revenue Model', 'Y17')))

def calc_Segment_Revenue_Model_Y39(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Y10')), xl_ref(ctx.cell('Segment Revenue Model', 'Y17')))

def calc_Segment_Revenue_Model_Y40(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Y11')), xl_ref(ctx.cell('Segment Revenue Model', 'Y17')))

def calc_Segment_Revenue_Model_Y41(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Y12')), xl_ref(ctx.cell('Segment Revenue Model', 'Y17')))

def calc_Segment_Revenue_Model_Y42(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Y13')), xl_ref(ctx.cell('Segment Revenue Model', 'Y17')))

def calc_Segment_Revenue_Model_Y43(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Y14')), xl_ref(ctx.cell('Segment Revenue Model', 'Y17')))

def calc_Segment_Revenue_Model_Y44(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Y15')), xl_ref(ctx.cell('Segment Revenue Model', 'Y17')))

def calc_Segment_Revenue_Model_Y45(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Y16')), xl_ref(ctx.cell('Segment Revenue Model', 'Y17')))

def calc_Segment_Revenue_Model_Y46(ctx):
    return xl_div(xl_ref(ctx.cell('Segment Revenue Model', 'Y17')), xl_ref(ctx.cell('Segment Revenue Model', 'Y17')))

def calc_INCOME_STATEMENT_H36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'H33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'H32'))))

def calc_INCOME_STATEMENT_H37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'H33')), xl_ref(ctx.cell('INCOME STATEMENT', 'H34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'H35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'H32'))), 100)

def calc_INCOME_STATEMENT_H38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'H32')), xl_ref(ctx.cell('INCOME STATEMENT', 'H33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'H34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'H35')))

def calc_CASH_FOW_STATEMENT_B8(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'H32'))

def calc_Ratio_Analysis_C12(ctx):
    return xl_div(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'H32')), xl_ref(ctx.cell('INCOME STATEMENT', 'H26'))), xl_mul(xl_add(xl_ref(ctx.cell('BALANCESHEET', 'C24')), xl_ref(ctx.cell('BALANCESHEET', 'B24'))), 0.5))

def calc_Ratio_Analysis_C18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'H29')), xl_ref(ctx.cell('INCOME STATEMENT', 'H32')))

def calc_INCOME_STATEMENT_O36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'O33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'O32'))))

def calc_INCOME_STATEMENT_O37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'O33')), xl_ref(ctx.cell('INCOME STATEMENT', 'O34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'O35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'O32'))), 100)

def calc_INCOME_STATEMENT_O38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'O32')), xl_ref(ctx.cell('INCOME STATEMENT', 'O33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'O34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'O35')))

def calc_CASH_FOW_STATEMENT_C8(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'O32'))

def calc_Ratio_Analysis_D12(ctx):
    return xl_div(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'O32')), xl_ref(ctx.cell('INCOME STATEMENT', 'O26'))), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'D8')), xl_ref(ctx.cell('BALANCESHEET', 'C8'))))

def calc_Ratio_Analysis_D18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'O29')), xl_ref(ctx.cell('INCOME STATEMENT', 'O32')))

def calc_INCOME_STATEMENT_D45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'D41')), xl_ref(ctx.cell('INCOME STATEMENT', 'D6'))), 100)

def calc_INCOME_STATEMENT_D49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'D41')), xl_ref(ctx.cell('INCOME STATEMENT', 'D48')))

def calc_INCOME_STATEMENT_F41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'F38')), xl_ref(ctx.cell('INCOME STATEMENT', 'F39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'F40')))

def calc_INCOME_STATEMENT_F43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'F38')), xl_error('#REF!')), 1), 100)

def calc_INCOME_STATEMENT_M41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'M38')), xl_ref(ctx.cell('INCOME STATEMENT', 'M39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'M40')))

def calc_INCOME_STATEMENT_M43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'M38')), xl_ref(ctx.cell('INCOME STATEMENT', 'F38'))), 1), 100)

def calc_INCOME_STATEMENT_R45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'R41')), xl_ref(ctx.cell('INCOME STATEMENT', 'R6'))), 100)

def calc_INCOME_STATEMENT_T41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'T38')), xl_ref(ctx.cell('INCOME STATEMENT', 'T39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'T40')))

def calc_INCOME_STATEMENT_T43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'T38')), xl_ref(ctx.cell('INCOME STATEMENT', 'M38'))), 1), 100)

def calc_INCOME_STATEMENT_V36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'V33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'V32'))))

def calc_INCOME_STATEMENT_V37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'V33')), xl_ref(ctx.cell('INCOME STATEMENT', 'V34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'V35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'V32'))), 100)

def calc_INCOME_STATEMENT_V38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'V32')), xl_ref(ctx.cell('INCOME STATEMENT', 'V33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'V34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'V35')))

def calc_CASH_FOW_STATEMENT_D8(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'V32'))

def calc_Ratio_Analysis_E12(ctx):
    return xl_div(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'V32')), xl_ref(ctx.cell('INCOME STATEMENT', 'V26'))), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'E8')), xl_ref(ctx.cell('BALANCESHEET', 'D8'))))

def calc_Ratio_Analysis_E18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'V29')), xl_ref(ctx.cell('INCOME STATEMENT', 'V32')))

def calc_Valuation_B28(ctx):
    return xl_add(xl_ref(ctx.cell('Valuation', 'B26')), xl_ref(ctx.cell('Valuation', 'B27')))

def calc_INCOME_STATEMENT_Y45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Y41')), xl_ref(ctx.cell('INCOME STATEMENT', 'Y6'))), 100)

def calc_INCOME_STATEMENT_Y49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Y41')), xl_ref(ctx.cell('INCOME STATEMENT', 'Y48')))

def calc_INCOME_STATEMENT_AA41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AA38')), xl_ref(ctx.cell('INCOME STATEMENT', 'AA39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AA40')))

def calc_INCOME_STATEMENT_AA43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AA38')), xl_ref(ctx.cell('INCOME STATEMENT', 'T38'))), 1), 100)

def calc_INCOME_STATEMENT_AC36(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AC33')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AC32'))))

def calc_INCOME_STATEMENT_AC37(ctx):
    return xl_mul(xl_div(xl_add(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AC33')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AC35'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AC32'))), 100)

def calc_INCOME_STATEMENT_AC38(ctx):
    return xl_sub(xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'AC32')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC33'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AC34'))), xl_ref(ctx.cell('INCOME STATEMENT', 'AC35')))

def calc_CASH_FOW_STATEMENT_E8(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AC32'))

def calc_Ratio_Analysis_F12(ctx):
    return xl_div(xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AC32')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC26'))), xl_add(xl_ref(ctx.cell('BALANCESHEET', 'F8')), xl_ref(ctx.cell('BALANCESHEET', 'E8'))))

def calc_Ratio_Analysis_F18(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AC29')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC32')))

def calc_Valuation_C26(ctx):
    return xl_sub(xl_add(xl_ref(ctx.cell('Valuation', 'C23')), xl_ref(ctx.cell('Valuation', 'C24'))), xl_ref(ctx.cell('Valuation', 'C25')))

def calc_Ratio_Analysis_G25(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AD36')), 100)

def calc_Ratio_Analysis_G16(ctx):
    return xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AD37')))

def calc_Ratio_Analysis_G11(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AD38')), xl_ref(ctx.cell('BALANCESHEET', 'G8')))

def calc_Valuation_D23(ctx):
    return xl_sub(xl_ref(ctx.cell('Valuation', 'D21')), xl_ref(ctx.cell('Valuation', 'D22')))

def calc_Valuation_D48(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'D22')), xl_ref(ctx.cell('Valuation', 'D21')))

def calc_Ratio_Analysis_H25(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AE36')), 100)

def calc_Ratio_Analysis_H16(ctx):
    return xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AE37')))

def calc_Ratio_Analysis_H11(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AE38')), xl_ref(ctx.cell('BALANCESHEET', 'H8')))

def calc_CASH_FOW_STATEMENT_F10(ctx):
    return xl_sub(xl_ref(ctx.cell('CASH FOW STATEMENT', 'F8')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'F9')))

def calc_Valuation_E23(ctx):
    return xl_sub(xl_ref(ctx.cell('Valuation', 'E21')), xl_ref(ctx.cell('Valuation', 'E22')))

def calc_Valuation_E48(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'E22')), xl_ref(ctx.cell('Valuation', 'E21')))

def calc_Valuation_H44(ctx):
    return xl_ref(ctx.cell('Valuation', 'G44'))

def calc_Ratio_Analysis_I25(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AF36')), 100)

def calc_Ratio_Analysis_I16(ctx):
    return xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AF37')))

def calc_Ratio_Analysis_I11(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AF38')), xl_ref(ctx.cell('BALANCESHEET', 'I8')))

def calc_CASH_FOW_STATEMENT_G10(ctx):
    return xl_sub(xl_ref(ctx.cell('CASH FOW STATEMENT', 'G8')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'G9')))

def calc_Valuation_F23(ctx):
    return xl_sub(xl_ref(ctx.cell('Valuation', 'F21')), xl_ref(ctx.cell('Valuation', 'F22')))

def calc_Ratio_Analysis_J25(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AG36')), 100)

def calc_Ratio_Analysis_J16(ctx):
    return xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AG37')))

def calc_Ratio_Analysis_J11(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AG38')), xl_ref(ctx.cell('BALANCESHEET', 'J8')))

def calc_CASH_FOW_STATEMENT_H10(ctx):
    return xl_sub(xl_ref(ctx.cell('CASH FOW STATEMENT', 'H8')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'H9')))

def calc_Valuation_G23(ctx):
    return xl_sub(xl_ref(ctx.cell('Valuation', 'G21')), xl_ref(ctx.cell('Valuation', 'G22')))

def calc_INCOME_STATEMENT_K47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'J47'))

def calc_INCOME_STATEMENT_J48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'J47')), 10)

def calc_INCOME_STATEMENT_I49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'I41')), xl_ref(ctx.cell('INCOME STATEMENT', 'I48')))

def calc_INCOME_STATEMENT_AB47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AA47'))

def calc_INCOME_STATEMENT_AA48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AA47')), 10)

def calc_INCOME_STATEMENT_Z49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Z41')), xl_ref(ctx.cell('INCOME STATEMENT', 'Z48')))

def calc_Valuation_K46(ctx):
    return xl_ref(ctx.cell('Valuation', 'J46'))

def calc_Valuation_K19(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'J19')), xl_add(1, xl_ref(ctx.cell('Valuation', 'K37'))))

def calc_Valuation_J24(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'J45')), xl_ref(ctx.cell('Valuation', 'J19')))

def calc_Valuation_J27(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'J46')), xl_uminus(xl_ref(ctx.cell('Valuation', 'J19'))))

def calc_Valuation_I40(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'I24')), xl_ref(ctx.cell('Valuation', 'H24'))), 1)

def calc_PRESENTATION_H45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'H44'))

def calc_PRESENTATION_H47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'H37')), xl_ref(ctx.cell('PRESENTATION', 'H44')))

def calc_PRESENTATION_O45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'O44'))

def calc_PRESENTATION_O47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'O37')), xl_ref(ctx.cell('PRESENTATION', 'O44')))

def calc_PRESENTATION_V45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'V44'))

def calc_PRESENTATION_V47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'V37')), xl_ref(ctx.cell('PRESENTATION', 'V44')))

def calc_PRESENTATION_AC45(ctx):
    return xl_ref(ctx.cell('PRESENTATION', 'AC44'))

def calc_PRESENTATION_AC47(ctx):
    return xl_div(xl_ref(ctx.cell('PRESENTATION', 'AC37')), xl_ref(ctx.cell('PRESENTATION', 'AC44')))

def calc_Ratio_Analysis_C25(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'H36')), 100)

def calc_Ratio_Analysis_C16(ctx):
    return xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'H37')))

def calc_INCOME_STATEMENT_H41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'H38')), xl_ref(ctx.cell('INCOME STATEMENT', 'H39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'H40')))

def calc_INCOME_STATEMENT_B42(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'B38')), xl_ref(ctx.cell('INCOME STATEMENT', 'H38'))), 100)

def calc_INCOME_STATEMENT_C42(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'C38')), xl_ref(ctx.cell('INCOME STATEMENT', 'H38'))), 100)

def calc_INCOME_STATEMENT_D42(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'D38')), xl_ref(ctx.cell('INCOME STATEMENT', 'H38'))), 100)

def calc_INCOME_STATEMENT_E42(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'E38')), xl_ref(ctx.cell('INCOME STATEMENT', 'H38'))), 100)

def calc_INCOME_STATEMENT_F42(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'F38')), xl_ref(ctx.cell('INCOME STATEMENT', 'H38'))), 100)

def calc_INCOME_STATEMENT_G42(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'G38')), xl_ref(ctx.cell('INCOME STATEMENT', 'H38'))), 100)

def calc_INCOME_STATEMENT_H43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'H38')), xl_error('#REF!')), 1), 100)

def calc_Ratio_Analysis_C11(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'H38')), xl_ref(ctx.cell('BALANCESHEET', 'C8')))

def calc_CASH_FOW_STATEMENT_B10(ctx):
    return xl_sub(xl_ref(ctx.cell('CASH FOW STATEMENT', 'B8')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'B9')))

def calc_Ratio_Analysis_D25(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'O36')), 100)

def calc_Ratio_Analysis_D16(ctx):
    return xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'O37')))

def calc_INCOME_STATEMENT_O41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'O38')), xl_ref(ctx.cell('INCOME STATEMENT', 'O39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'O40')))

def calc_INCOME_STATEMENT_O43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'O38')), xl_ref(ctx.cell('INCOME STATEMENT', 'H38'))), 1), 100)

def calc_Ratio_Analysis_D11(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'O38')), xl_ref(ctx.cell('BALANCESHEET', 'D8')))

def calc_CASH_FOW_STATEMENT_C10(ctx):
    return xl_sub(xl_ref(ctx.cell('CASH FOW STATEMENT', 'C8')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'C9')))

def calc_INCOME_STATEMENT_F45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'F41')), xl_ref(ctx.cell('INCOME STATEMENT', 'F6'))), 100)

def calc_INCOME_STATEMENT_F49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'F41')), xl_ref(ctx.cell('INCOME STATEMENT', 'F48')))

def calc_INCOME_STATEMENT_M45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'M41')), xl_ref(ctx.cell('INCOME STATEMENT', 'M6'))), 100)

def calc_INCOME_STATEMENT_T45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'T41')), xl_ref(ctx.cell('INCOME STATEMENT', 'T6'))), 100)

def calc_INCOME_STATEMENT_T49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'T41')), xl_ref(ctx.cell('INCOME STATEMENT', 'T48')))

def calc_Ratio_Analysis_E25(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'V36')), 100)

def calc_Valuation_K10(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'V37'))

def calc_Ratio_Analysis_E16(ctx):
    return xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'V37')))

def calc_INCOME_STATEMENT_V41(ctx):
    return xl_sub(xl_sub(xl_ref(ctx.cell('INCOME STATEMENT', 'V38')), xl_ref(ctx.cell('INCOME STATEMENT', 'V39'))), xl_ref(ctx.cell('INCOME STATEMENT', 'V40')))

def calc_INCOME_STATEMENT_P42(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'P38')), xl_ref(ctx.cell('INCOME STATEMENT', 'V38'))), 100)

def calc_INCOME_STATEMENT_Q42(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Q38')), xl_ref(ctx.cell('INCOME STATEMENT', 'V38'))), 100)

def calc_INCOME_STATEMENT_R42(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'R38')), xl_ref(ctx.cell('INCOME STATEMENT', 'V38'))), 100)

def calc_INCOME_STATEMENT_S42(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'S38')), xl_ref(ctx.cell('INCOME STATEMENT', 'V38'))), 100)

def calc_INCOME_STATEMENT_T42(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'T38')), xl_ref(ctx.cell('INCOME STATEMENT', 'V38'))), 100)

def calc_INCOME_STATEMENT_U42(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'U38')), xl_ref(ctx.cell('INCOME STATEMENT', 'V38'))), 100)

def calc_INCOME_STATEMENT_V43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'V38')), xl_ref(ctx.cell('INCOME STATEMENT', 'O38'))), 1), 100)

def calc_Ratio_Analysis_E11(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'V38')), xl_ref(ctx.cell('BALANCESHEET', 'E8')))

def calc_CASH_FOW_STATEMENT_D10(ctx):
    return xl_sub(xl_ref(ctx.cell('CASH FOW STATEMENT', 'D8')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'D9')))

def calc_INCOME_STATEMENT_AC41(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AA41')), xl_ref(ctx.cell('INCOME STATEMENT', 'AB41')))

def calc_INCOME_STATEMENT_AA45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AA41')), xl_ref(ctx.cell('INCOME STATEMENT', 'AA6'))), 100)

def calc_Ratio_Analysis_F25(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AC36')), 100)

def calc_Ratio_Analysis_F16(ctx):
    return xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AC37')))

def calc_INCOME_STATEMENT_AC43(ctx):
    return xl_mul(xl_sub(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AC38')), xl_ref(ctx.cell('INCOME STATEMENT', 'V38'))), 1), 100)

def calc_Ratio_Analysis_F11(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AC38')), xl_ref(ctx.cell('BALANCESHEET', 'F8')))

def calc_CASH_FOW_STATEMENT_E10(ctx):
    return xl_sub(xl_ref(ctx.cell('CASH FOW STATEMENT', 'E8')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'E9')))

def calc_Valuation_C28(ctx):
    return xl_add(xl_ref(ctx.cell('Valuation', 'C26')), xl_ref(ctx.cell('Valuation', 'C27')))

def calc_Valuation_D26(ctx):
    return xl_sub(xl_add(xl_ref(ctx.cell('Valuation', 'D23')), xl_ref(ctx.cell('Valuation', 'D24'))), xl_ref(ctx.cell('Valuation', 'D25')))

def calc_CASH_FOW_STATEMENT_F12(ctx):
    return xl_add(xl_ref(ctx.cell('CASH FOW STATEMENT', 'F10')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'F11')))

def calc_Valuation_E26(ctx):
    return xl_sub(xl_add(xl_ref(ctx.cell('Valuation', 'E23')), xl_ref(ctx.cell('Valuation', 'E24'))), xl_ref(ctx.cell('Valuation', 'E25')))

def calc_Valuation_F48(ctx):
    return xl_ref(ctx.cell('Valuation', 'E48'))

def calc_Valuation_H21(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'H19')), xl_ref(ctx.cell('Valuation', 'H44')))

def calc_Valuation_I44(ctx):
    return xl_ref(ctx.cell('Valuation', 'H44'))

def calc_CASH_FOW_STATEMENT_G12(ctx):
    return xl_add(xl_ref(ctx.cell('CASH FOW STATEMENT', 'G10')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'G11')))

def calc_Valuation_F26(ctx):
    return xl_sub(xl_add(xl_ref(ctx.cell('Valuation', 'F23')), xl_ref(ctx.cell('Valuation', 'F24'))), xl_ref(ctx.cell('Valuation', 'F25')))

def calc_CASH_FOW_STATEMENT_H12(ctx):
    return xl_add(xl_ref(ctx.cell('CASH FOW STATEMENT', 'H10')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'H11')))

def calc_Valuation_G26(ctx):
    return xl_sub(xl_add(xl_ref(ctx.cell('Valuation', 'G23')), xl_ref(ctx.cell('Valuation', 'G24'))), xl_ref(ctx.cell('Valuation', 'G25')))

def calc_INCOME_STATEMENT_L47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'K47'))

def calc_INCOME_STATEMENT_K48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'K47')), 10)

def calc_INCOME_STATEMENT_J49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'J41')), xl_ref(ctx.cell('INCOME STATEMENT', 'J48')))

def calc_INCOME_STATEMENT_AC47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AB47'))

def calc_INCOME_STATEMENT_AB48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AB47')), 10)

def calc_INCOME_STATEMENT_AA49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AA41')), xl_ref(ctx.cell('INCOME STATEMENT', 'AA48')))

def calc_Valuation_K24(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'K45')), xl_ref(ctx.cell('Valuation', 'K19')))

def calc_Valuation_K27(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'K46')), xl_uminus(xl_ref(ctx.cell('Valuation', 'K19'))))

def calc_Valuation_J40(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'J24')), xl_ref(ctx.cell('Valuation', 'I24'))), 1)

def calc_INCOME_STATEMENT_H45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'H41')), xl_ref(ctx.cell('INCOME STATEMENT', 'H6'))), 100)

def calc_INCOME_STATEMENT_H49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'H41')), xl_ref(ctx.cell('INCOME STATEMENT', 'H48')))

def calc_Ratio_Analysis_C14(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'H41')), 100), xl_ref(ctx.cell('INCOME STATEMENT', 'H6')))

def calc_INCOME_STATEMENT_H42(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'F42')), xl_ref(ctx.cell('INCOME STATEMENT', 'G42')))

def calc_CASH_FOW_STATEMENT_B12(ctx):
    return xl_add(xl_ref(ctx.cell('CASH FOW STATEMENT', 'B10')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'B11')))

def calc_COMPANY_OVERVIEW_C23(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'O41')), xl_ref(ctx.cell('INCOME STATEMENT', 'O6')))

def calc_INCOME_STATEMENT_I42(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'I41')), xl_ref(ctx.cell('INCOME STATEMENT', 'O41'))), 100)

def calc_INCOME_STATEMENT_J42(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'J41')), xl_ref(ctx.cell('INCOME STATEMENT', 'O41'))), 100)

def calc_INCOME_STATEMENT_K42(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'K41')), xl_ref(ctx.cell('INCOME STATEMENT', 'O41'))), 100)

def calc_INCOME_STATEMENT_L42(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'L41')), xl_ref(ctx.cell('INCOME STATEMENT', 'O41'))), 100)

def calc_INCOME_STATEMENT_M42(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'M41')), xl_ref(ctx.cell('INCOME STATEMENT', 'O41'))), 100)

def calc_INCOME_STATEMENT_N42(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'N41')), xl_ref(ctx.cell('INCOME STATEMENT', 'O41'))), 100)

def calc_INCOME_STATEMENT_O45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'O41')), xl_ref(ctx.cell('INCOME STATEMENT', 'O6'))), 100)

def calc_Ratio_Analysis_D14(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'O41')), 100), xl_ref(ctx.cell('INCOME STATEMENT', 'O6')))

def calc_CASH_FOW_STATEMENT_C12(ctx):
    return xl_add(xl_ref(ctx.cell('CASH FOW STATEMENT', 'C10')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'C11')))

def calc_Valuation_K13(ctx):
    return xl_add(xl_mul(xl_ref(ctx.cell('Valuation', 'B10')), xl_ref(ctx.cell('Valuation', 'K6'))), xl_mul(xl_mul(xl_ref(ctx.cell('Valuation', 'K9')), xl_ref(ctx.cell('Valuation', 'K7'))), xl_sub(1, xl_ref(ctx.cell('Valuation', 'K10')))))

def calc_COMPANY_OVERVIEW_D23(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'V41')), xl_ref(ctx.cell('INCOME STATEMENT', 'V6')))

def calc_COMPANY_OVERVIEW_E23(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'V41')), xl_ref(ctx.cell('INCOME STATEMENT', 'V6')))

def calc_INCOME_STATEMENT_V45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'V41')), xl_ref(ctx.cell('INCOME STATEMENT', 'V6'))), 100)

def calc_INCOME_STATEMENT_V49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'V41')), xl_ref(ctx.cell('INCOME STATEMENT', 'V48')))

def calc_Ratio_Analysis_E14(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'V41')), 100), xl_ref(ctx.cell('INCOME STATEMENT', 'V6')))

def calc_INCOME_STATEMENT_V42(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'T42')), xl_ref(ctx.cell('INCOME STATEMENT', 'U42')))

def calc_CASH_FOW_STATEMENT_D12(ctx):
    return xl_add(xl_ref(ctx.cell('CASH FOW STATEMENT', 'D10')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'D11')))

def calc_COMPANY_OVERVIEW_F23(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AC41')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC6')))

def calc_INCOME_STATEMENT_W42(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'W41')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC41'))), 100)

def calc_INCOME_STATEMENT_X42(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'X41')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC41'))), 100)

def calc_INCOME_STATEMENT_Y42(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Y41')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC41'))), 100)

def calc_INCOME_STATEMENT_Z42(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Z41')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC41'))), 100)

def calc_INCOME_STATEMENT_AA42(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AA41')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC41'))), 100)

def calc_INCOME_STATEMENT_AB42(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AB41')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC41'))), 100)

def calc_INCOME_STATEMENT_AD43(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AD41')), xl_percent(xl_ref(ctx.cell('INCOME STATEMENT', 'AC41'))))

def calc_INCOME_STATEMENT_AC45(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AC41')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC6'))), 100)

def calc_Ratio_Analysis_F14(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('INCOME STATEMENT', 'AC41')), 100), xl_ref(ctx.cell('INCOME STATEMENT', 'AC6')))

def calc_CASH_FOW_STATEMENT_E12(ctx):
    return xl_add(xl_ref(ctx.cell('CASH FOW STATEMENT', 'E10')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'E11')))

def calc_Valuation_D28(ctx):
    return xl_add(xl_ref(ctx.cell('Valuation', 'D26')), xl_ref(ctx.cell('Valuation', 'D27')))

def calc_CASH_FOW_STATEMENT_F14(ctx):
    return xl_sub(xl_ref(ctx.cell('CASH FOW STATEMENT', 'F12')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'F13')))

def calc_Valuation_E28(ctx):
    return xl_add(xl_ref(ctx.cell('Valuation', 'E26')), xl_ref(ctx.cell('Valuation', 'E27')))

def calc_Valuation_G48(ctx):
    return xl_ref(ctx.cell('Valuation', 'F48'))

def calc_Valuation_H20(ctx):
    return xl_add(xl_ref(ctx.cell('Valuation', 'H21')), xl_ref(ctx.cell('Valuation', 'H24')))

def calc_Valuation_H39(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'H21')), xl_ref(ctx.cell('Valuation', 'G21'))), 1)

def calc_Valuation_I21(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'I19')), xl_ref(ctx.cell('Valuation', 'I44')))

def calc_Valuation_J44(ctx):
    return xl_ref(ctx.cell('Valuation', 'I44'))

def calc_CASH_FOW_STATEMENT_G14(ctx):
    return xl_sub(xl_ref(ctx.cell('CASH FOW STATEMENT', 'G12')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'G13')))

def calc_Valuation_F28(ctx):
    return xl_add(xl_ref(ctx.cell('Valuation', 'F26')), xl_ref(ctx.cell('Valuation', 'F27')))

def calc_CASH_FOW_STATEMENT_H14(ctx):
    return xl_sub(xl_ref(ctx.cell('CASH FOW STATEMENT', 'H12')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'H13')))

def calc_Valuation_G28(ctx):
    return xl_add(xl_ref(ctx.cell('Valuation', 'G26')), xl_ref(ctx.cell('Valuation', 'G27')))

def calc_INCOME_STATEMENT_M47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'L47'))

def calc_INCOME_STATEMENT_L48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'L47')), 10)

def calc_INCOME_STATEMENT_K49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'K41')), xl_ref(ctx.cell('INCOME STATEMENT', 'K48')))

def calc_INCOME_STATEMENT_AD47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AC47'))

def calc_INCOME_STATEMENT_AC48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AC47')), 10)

def calc_INCOME_STATEMENT_AB49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AB41')), xl_ref(ctx.cell('INCOME STATEMENT', 'AB48')))

def calc_Valuation_K40(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'K24')), xl_ref(ctx.cell('Valuation', 'J24'))), 1)

def calc_COMPANY_OVERVIEW_C21(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'H49'))

def calc_Ratio_Analysis_C27(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'H49'))

def calc_PRESENTATION_C65(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'C14'))

def calc_CASH_FOW_STATEMENT_B14(ctx):
    return xl_sub(xl_ref(ctx.cell('CASH FOW STATEMENT', 'B12')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'B13')))

def calc_INCOME_STATEMENT_O42(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'M42')), xl_ref(ctx.cell('INCOME STATEMENT', 'N42')))

def calc_PRESENTATION_D65(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'D14'))

def calc_CASH_FOW_STATEMENT_C14(ctx):
    return xl_sub(xl_ref(ctx.cell('CASH FOW STATEMENT', 'C12')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'C13')))

def calc_Valuation_C33(ctx):
    return xl_div(1, xl_pow(xl_add(1, xl_ref(ctx.cell('Valuation', 'K13'))), xl_ref(ctx.cell('Valuation', 'C32'))))

def calc_Valuation_D33(ctx):
    return xl_div(1, xl_pow(xl_add(1, xl_ref(ctx.cell('Valuation', 'K13'))), xl_ref(ctx.cell('Valuation', 'D32'))))

def calc_Valuation_E33(ctx):
    return xl_div(1, xl_pow(xl_add(1, xl_ref(ctx.cell('Valuation', 'K13'))), xl_ref(ctx.cell('Valuation', 'E32'))))

def calc_Valuation_F33(ctx):
    return xl_div(1, xl_pow(xl_add(1, xl_ref(ctx.cell('Valuation', 'K13'))), xl_ref(ctx.cell('Valuation', 'F32'))))

def calc_Valuation_G33(ctx):
    return xl_div(1, xl_pow(xl_add(1, xl_ref(ctx.cell('Valuation', 'K13'))), xl_ref(ctx.cell('Valuation', 'G32'))))

def calc_Valuation_H33(ctx):
    return xl_div(1, xl_pow(xl_add(1, xl_ref(ctx.cell('Valuation', 'K13'))), xl_ref(ctx.cell('Valuation', 'H32'))))

def calc_Valuation_I33(ctx):
    return xl_div(1, xl_pow(xl_add(1, xl_ref(ctx.cell('Valuation', 'K13'))), xl_ref(ctx.cell('Valuation', 'I32'))))

def calc_Valuation_J33(ctx):
    return xl_div(1, xl_pow(xl_add(1, xl_ref(ctx.cell('Valuation', 'K13'))), xl_ref(ctx.cell('Valuation', 'J32'))))

def calc_Valuation_B53(ctx):
    return xl_ref(ctx.cell('Valuation', 'K13'))

def calc_COMPANY_OVERVIEW_E21(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'V49'))

def calc_COMPANY_OVERVIEW_E27(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'G7')), xl_ref(ctx.cell('INCOME STATEMENT', 'V49')))

def calc_Ratio_Analysis_E27(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'V49'))

def calc_PRESENTATION_E65(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'E14'))

def calc_CASH_FOW_STATEMENT_D14(ctx):
    return xl_sub(xl_ref(ctx.cell('CASH FOW STATEMENT', 'D12')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'D13')))

def calc_INCOME_STATEMENT_AC42(ctx):
    return xl_add(xl_ref(ctx.cell('INCOME STATEMENT', 'AA42')), xl_ref(ctx.cell('INCOME STATEMENT', 'AB42')))

def calc_PRESENTATION_F65(ctx):
    return xl_ref(ctx.cell('Ratio Analysis', 'F14'))

def calc_CASH_FOW_STATEMENT_E14(ctx):
    return xl_sub(xl_ref(ctx.cell('CASH FOW STATEMENT', 'E12')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'E13')))

def calc_CASH_FOW_STATEMENT_F23(ctx):
    return xl_add(xl_add(xl_sub(xl_sub(xl_add(xl_sub(xl_ref(ctx.cell('CASH FOW STATEMENT', 'F14')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'F16'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'F17'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'F19'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'E20'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'F18'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'F21')))

def calc_Valuation_H48(ctx):
    return xl_ref(ctx.cell('Valuation', 'G48'))

def calc_Valuation_H38(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'H20')), xl_ref(ctx.cell('Valuation', 'G20'))), 1)

def calc_Valuation_I20(ctx):
    return xl_add(xl_ref(ctx.cell('Valuation', 'I21')), xl_ref(ctx.cell('Valuation', 'I24')))

def calc_Valuation_I39(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'I21')), xl_ref(ctx.cell('Valuation', 'H21'))), 1)

def calc_Valuation_J21(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'J19')), xl_ref(ctx.cell('Valuation', 'J44')))

def calc_Valuation_K44(ctx):
    return xl_ref(ctx.cell('Valuation', 'J44'))

def calc_CASH_FOW_STATEMENT_G23(ctx):
    return xl_add(xl_add(xl_sub(xl_sub(xl_add(xl_sub(xl_ref(ctx.cell('CASH FOW STATEMENT', 'G14')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'G16'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'G17'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'G19'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'F20'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'G18'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'G21')))

def calc_CASH_FOW_STATEMENT_H23(ctx):
    return xl_add(xl_add(xl_sub(xl_sub(xl_add(xl_sub(xl_ref(ctx.cell('CASH FOW STATEMENT', 'H14')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'H16'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'H17'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'H19'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'G20'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'H18'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'H21')))

def calc_INCOME_STATEMENT_N47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'M47'))

def calc_INCOME_STATEMENT_M48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'M47')), 1)

def calc_INCOME_STATEMENT_L49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'L41')), xl_ref(ctx.cell('INCOME STATEMENT', 'L48')))

def calc_INCOME_STATEMENT_AE47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AD47'))

def calc_INCOME_STATEMENT_AD48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AD47')), 10)

def calc_INCOME_STATEMENT_AC49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AC41')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC48')))

def calc_Valuation_C30(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AC48'))

def calc_CASH_FOW_STATEMENT_B23(ctx):
    return xl_add(xl_add(xl_sub(xl_sub(xl_add(xl_sub(xl_ref(ctx.cell('CASH FOW STATEMENT', 'B14')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'B16'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'B17'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'B19'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'B20'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'B18'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'B21')))

def calc_CASH_FOW_STATEMENT_C23(ctx):
    return xl_sub(xl_add(xl_sub(xl_ref(ctx.cell('CASH FOW STATEMENT', 'C14')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'C16'))), xl_sub(xl_ref(ctx.cell('CASH FOW STATEMENT', 'C17')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'C19')))), xl_add(xl_add(xl_ref(ctx.cell('CASH FOW STATEMENT', 'B20')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'C18'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'C21'))))

def calc_Valuation_C34(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'C33')), xl_ref(ctx.cell('Valuation', 'C28')))

def calc_Valuation_D34(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'D33')), xl_ref(ctx.cell('Valuation', 'D28')))

def calc_Valuation_E34(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'E33')), xl_ref(ctx.cell('Valuation', 'E28')))

def calc_Valuation_F34(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'F33')), xl_ref(ctx.cell('Valuation', 'F28')))

def calc_Valuation_G34(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'G33')), xl_ref(ctx.cell('Valuation', 'G28')))

def calc_Valuation_A67(ctx):
    return xl_ref(ctx.cell('Valuation', 'B53'))

def calc_Valuation_A77(ctx):
    return xl_ref(ctx.cell('Valuation', 'B53'))

def calc_CASH_FOW_STATEMENT_D23(ctx):
    return xl_sub(xl_add(xl_sub(xl_ref(ctx.cell('CASH FOW STATEMENT', 'D14')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'D16'))), xl_sub(xl_ref(ctx.cell('CASH FOW STATEMENT', 'D17')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'D19')))), xl_add(xl_add(xl_ref(ctx.cell('CASH FOW STATEMENT', 'C20')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'D18'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'D21'))))

def calc_CASH_FOW_STATEMENT_E23(ctx):
    return xl_add(xl_add(xl_sub(xl_sub(xl_add(xl_sub(xl_ref(ctx.cell('CASH FOW STATEMENT', 'E14')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'E16'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'E17'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'E19'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'D20'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'E18'))), xl_ref(ctx.cell('CASH FOW STATEMENT', 'E21')))

def calc_CASH_FOW_STATEMENT_F24(ctx):
    return xl_add(xl_ref(ctx.cell('CASH FOW STATEMENT', 'F22')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'F23')))

def calc_Valuation_H22(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'H21')), xl_ref(ctx.cell('Valuation', 'H48')))

def calc_Valuation_I48(ctx):
    return xl_ref(ctx.cell('Valuation', 'H48'))

def calc_Valuation_I38(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'I20')), xl_ref(ctx.cell('Valuation', 'H20'))), 1)

def calc_Valuation_J20(ctx):
    return xl_add(xl_ref(ctx.cell('Valuation', 'J21')), xl_ref(ctx.cell('Valuation', 'J24')))

def calc_Valuation_J39(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'J21')), xl_ref(ctx.cell('Valuation', 'I21'))), 1)

def calc_Valuation_K21(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'K19')), xl_ref(ctx.cell('Valuation', 'K44')))

def calc_CASH_FOW_STATEMENT_G24(ctx):
    return xl_add(xl_ref(ctx.cell('CASH FOW STATEMENT', 'G22')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'G23')))

def calc_CASH_FOW_STATEMENT_H24(ctx):
    return xl_add(xl_ref(ctx.cell('CASH FOW STATEMENT', 'H22')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'H23')))

def calc_INCOME_STATEMENT_O47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'N47'))

def calc_INCOME_STATEMENT_P47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'N47'))

def calc_INCOME_STATEMENT_N48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'N47')), 10)

def calc_INCOME_STATEMENT_M49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'M41')), xl_ref(ctx.cell('INCOME STATEMENT', 'M48')))

def calc_INCOME_STATEMENT_AF47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AE47'))

def calc_INCOME_STATEMENT_AE48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AE47')), 10)

def calc_INCOME_STATEMENT_AD49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AD41')), xl_ref(ctx.cell('INCOME STATEMENT', 'AD48')))

def calc_Valuation_D30(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AD48'))

def calc_COMPANY_OVERVIEW_F21(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AC49'))

def calc_COMPANY_OVERVIEW_F27(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'G7')), xl_ref(ctx.cell('INCOME STATEMENT', 'AC49')))

def calc_Ratio_Analysis_F27(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AC49'))

def calc_CASH_FOW_STATEMENT_E24(ctx):
    return xl_add(xl_ref(ctx.cell('CASH FOW STATEMENT', 'E22')), xl_ref(ctx.cell('CASH FOW STATEMENT', 'E23')))

def calc_Valuation_H23(ctx):
    return xl_sub(xl_ref(ctx.cell('Valuation', 'H21')), xl_ref(ctx.cell('Valuation', 'H22')))

def calc_Valuation_I22(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'I21')), xl_ref(ctx.cell('Valuation', 'I48')))

def calc_Valuation_J48(ctx):
    return xl_ref(ctx.cell('Valuation', 'I48'))

def calc_Valuation_J38(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'J20')), xl_ref(ctx.cell('Valuation', 'I20'))), 1)

def calc_Valuation_K20(ctx):
    return xl_add(xl_ref(ctx.cell('Valuation', 'K21')), xl_ref(ctx.cell('Valuation', 'K24')))

def calc_Valuation_K39(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'K21')), xl_ref(ctx.cell('Valuation', 'J21'))), 1)

def calc_Valuation_D75(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('Valuation', 'D74')), xl_add(xl_ref(ctx.cell('Valuation', 'K21')), xl_ref(ctx.cell('Valuation', 'K24')))), xl_pow(xl_add(1, xl_ref(ctx.cell('Valuation', 'A75'))), xl_ref(ctx.cell('Valuation', 'J32'))))

def calc_Valuation_E75(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('Valuation', 'E74')), xl_add(xl_ref(ctx.cell('Valuation', 'K21')), xl_ref(ctx.cell('Valuation', 'K24')))), xl_pow(xl_add(1, xl_ref(ctx.cell('Valuation', 'A75'))), xl_ref(ctx.cell('Valuation', 'J32'))))

def calc_Valuation_F75(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('Valuation', 'F74')), xl_add(xl_ref(ctx.cell('Valuation', 'K21')), xl_ref(ctx.cell('Valuation', 'K24')))), xl_pow(xl_add(1, xl_ref(ctx.cell('Valuation', 'A75'))), xl_ref(ctx.cell('Valuation', 'J32'))))

def calc_Valuation_D76(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('Valuation', 'D74')), xl_add(xl_ref(ctx.cell('Valuation', 'K21')), xl_ref(ctx.cell('Valuation', 'K24')))), xl_pow(xl_add(1, xl_ref(ctx.cell('Valuation', 'A76'))), xl_ref(ctx.cell('Valuation', 'J32'))))

def calc_Valuation_E76(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('Valuation', 'E74')), xl_add(xl_ref(ctx.cell('Valuation', 'K21')), xl_ref(ctx.cell('Valuation', 'K24')))), xl_pow(xl_add(1, xl_ref(ctx.cell('Valuation', 'A76'))), xl_ref(ctx.cell('Valuation', 'J32'))))

def calc_Valuation_F76(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('Valuation', 'F74')), xl_add(xl_ref(ctx.cell('Valuation', 'K21')), xl_ref(ctx.cell('Valuation', 'K24')))), xl_pow(xl_add(1, xl_ref(ctx.cell('Valuation', 'A76'))), xl_ref(ctx.cell('Valuation', 'J32'))))

def calc_Valuation_D77(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('Valuation', 'D74')), xl_add(xl_ref(ctx.cell('Valuation', 'K21')), xl_ref(ctx.cell('Valuation', 'K24')))), xl_pow(xl_add(1, xl_ref(ctx.cell('Valuation', 'A77'))), xl_ref(ctx.cell('Valuation', 'J32'))))

def calc_Valuation_E77(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('Valuation', 'E74')), xl_add(xl_ref(ctx.cell('Valuation', 'K21')), xl_ref(ctx.cell('Valuation', 'K24')))), xl_pow(xl_add(1, xl_ref(ctx.cell('Valuation', 'A77'))), xl_ref(ctx.cell('Valuation', 'J32'))))

def calc_Valuation_F77(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('Valuation', 'F74')), xl_add(xl_ref(ctx.cell('Valuation', 'K21')), xl_ref(ctx.cell('Valuation', 'K24')))), xl_pow(xl_add(1, xl_ref(ctx.cell('Valuation', 'A77'))), xl_ref(ctx.cell('Valuation', 'J32'))))

def calc_Valuation_D78(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('Valuation', 'D74')), xl_add(xl_ref(ctx.cell('Valuation', 'K21')), xl_ref(ctx.cell('Valuation', 'K24')))), xl_pow(xl_add(1, xl_ref(ctx.cell('Valuation', 'A78'))), xl_ref(ctx.cell('Valuation', 'J32'))))

def calc_Valuation_E78(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('Valuation', 'E74')), xl_add(xl_ref(ctx.cell('Valuation', 'K21')), xl_ref(ctx.cell('Valuation', 'K24')))), xl_pow(xl_add(1, xl_ref(ctx.cell('Valuation', 'A78'))), xl_ref(ctx.cell('Valuation', 'J32'))))

def calc_Valuation_F78(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('Valuation', 'F74')), xl_add(xl_ref(ctx.cell('Valuation', 'K21')), xl_ref(ctx.cell('Valuation', 'K24')))), xl_pow(xl_add(1, xl_ref(ctx.cell('Valuation', 'A78'))), xl_ref(ctx.cell('Valuation', 'J32'))))

def calc_Valuation_D79(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('Valuation', 'D74')), xl_add(xl_ref(ctx.cell('Valuation', 'K21')), xl_ref(ctx.cell('Valuation', 'K24')))), xl_pow(xl_add(1, xl_ref(ctx.cell('Valuation', 'A79'))), xl_ref(ctx.cell('Valuation', 'J32'))))

def calc_Valuation_E79(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('Valuation', 'E74')), xl_add(xl_ref(ctx.cell('Valuation', 'K21')), xl_ref(ctx.cell('Valuation', 'K24')))), xl_pow(xl_add(1, xl_ref(ctx.cell('Valuation', 'A79'))), xl_ref(ctx.cell('Valuation', 'J32'))))

def calc_Valuation_F79(ctx):
    return xl_div(xl_mul(xl_ref(ctx.cell('Valuation', 'F74')), xl_add(xl_ref(ctx.cell('Valuation', 'K21')), xl_ref(ctx.cell('Valuation', 'K24')))), xl_pow(xl_add(1, xl_ref(ctx.cell('Valuation', 'A79'))), xl_ref(ctx.cell('Valuation', 'J32'))))

def calc_INCOME_STATEMENT_O48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'O47')), 10)

def calc_INCOME_STATEMENT_Q47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'P47'))

def calc_INCOME_STATEMENT_P48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'P47')), 10)

def calc_INCOME_STATEMENT_N49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'N41')), xl_ref(ctx.cell('INCOME STATEMENT', 'N48')))

def calc_INCOME_STATEMENT_AG47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AF47'))

def calc_INCOME_STATEMENT_AF48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AF47')), 10)

def calc_INCOME_STATEMENT_AE49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AE41')), xl_ref(ctx.cell('INCOME STATEMENT', 'AE48')))

def calc_Valuation_E30(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AE48'))

def calc_COMPANY_OVERVIEW_G21(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AD49'))

def calc_COMPANY_OVERVIEW_G27(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'G7')), xl_ref(ctx.cell('INCOME STATEMENT', 'AD49')))

def calc_Ratio_Analysis_G27(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AD49'))

def calc_Valuation_H26(ctx):
    return xl_sub(xl_add(xl_ref(ctx.cell('Valuation', 'H23')), xl_ref(ctx.cell('Valuation', 'H24'))), xl_ref(ctx.cell('Valuation', 'H25')))

def calc_Valuation_I23(ctx):
    return xl_sub(xl_ref(ctx.cell('Valuation', 'I21')), xl_ref(ctx.cell('Valuation', 'I22')))

def calc_Valuation_J22(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'J21')), xl_ref(ctx.cell('Valuation', 'J48')))

def calc_Valuation_K48(ctx):
    return xl_ref(ctx.cell('Valuation', 'J48'))

def calc_Valuation_K38(ctx):
    return xl_sub(xl_div(xl_ref(ctx.cell('Valuation', 'K20')), xl_ref(ctx.cell('Valuation', 'J20'))), 1)

def calc_INCOME_STATEMENT_O49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'O41')), xl_ref(ctx.cell('INCOME STATEMENT', 'O48')))

def calc_INCOME_STATEMENT_R47(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'Q47'))

def calc_INCOME_STATEMENT_Q48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Q47')), 10)

def calc_INCOME_STATEMENT_P49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'P41')), xl_ref(ctx.cell('INCOME STATEMENT', 'P48')))

def calc_INCOME_STATEMENT_AG48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AG47')), 10)

def calc_INCOME_STATEMENT_AF49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AF41')), xl_ref(ctx.cell('INCOME STATEMENT', 'AF48')))

def calc_Valuation_F30(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AF48'))

def calc_COMPANY_OVERVIEW_H21(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AE49'))

def calc_COMPANY_OVERVIEW_H27(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'G7')), xl_ref(ctx.cell('INCOME STATEMENT', 'AE49')))

def calc_Ratio_Analysis_H27(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AE49'))

def calc_Valuation_H28(ctx):
    return xl_add(xl_ref(ctx.cell('Valuation', 'H26')), xl_ref(ctx.cell('Valuation', 'H27')))

def calc_Valuation_I26(ctx):
    return xl_sub(xl_add(xl_ref(ctx.cell('Valuation', 'I23')), xl_ref(ctx.cell('Valuation', 'I24'))), xl_ref(ctx.cell('Valuation', 'I25')))

def calc_Valuation_J23(ctx):
    return xl_sub(xl_ref(ctx.cell('Valuation', 'J21')), xl_ref(ctx.cell('Valuation', 'J22')))

def calc_Valuation_K22(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'K21')), xl_ref(ctx.cell('Valuation', 'K48')))

def calc_COMPANY_OVERVIEW_D21(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'O49'))

def calc_Ratio_Analysis_D27(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'O49'))

def calc_INCOME_STATEMENT_R48(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'R47')), 10)

def calc_INCOME_STATEMENT_Q49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'Q41')), xl_ref(ctx.cell('INCOME STATEMENT', 'Q48')))

def calc_INCOME_STATEMENT_AG49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'AG41')), xl_ref(ctx.cell('INCOME STATEMENT', 'AG48')))

def calc_Valuation_G30(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AG48'))

def calc_COMPANY_OVERVIEW_I21(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AF49'))

def calc_COMPANY_OVERVIEW_I27(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'G7')), xl_ref(ctx.cell('INCOME STATEMENT', 'AF49')))

def calc_Ratio_Analysis_I27(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AF49'))

def calc_Valuation_H34(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'H33')), xl_ref(ctx.cell('Valuation', 'H28')))

def calc_Valuation_I28(ctx):
    return xl_add(xl_ref(ctx.cell('Valuation', 'I26')), xl_ref(ctx.cell('Valuation', 'I27')))

def calc_Valuation_J26(ctx):
    return xl_sub(xl_add(xl_ref(ctx.cell('Valuation', 'J23')), xl_ref(ctx.cell('Valuation', 'J24'))), xl_ref(ctx.cell('Valuation', 'J25')))

def calc_Valuation_K23(ctx):
    return xl_sub(xl_ref(ctx.cell('Valuation', 'K21')), xl_ref(ctx.cell('Valuation', 'K22')))

def calc_INCOME_STATEMENT_R49(ctx):
    return xl_div(xl_ref(ctx.cell('INCOME STATEMENT', 'R41')), xl_ref(ctx.cell('INCOME STATEMENT', 'R48')))

def calc_COMPANY_OVERVIEW_J21(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AG49'))

def calc_COMPANY_OVERVIEW_J27(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'G7')), xl_ref(ctx.cell('INCOME STATEMENT', 'AG49')))

def calc_Ratio_Analysis_J27(ctx):
    return xl_ref(ctx.cell('INCOME STATEMENT', 'AG49'))

def calc_Valuation_H30(ctx):
    return xl_ref(ctx.cell('Valuation', 'G30'))

def calc_Valuation_I34(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'I33')), xl_ref(ctx.cell('Valuation', 'I28')))

def calc_Valuation_J28(ctx):
    return xl_add(xl_ref(ctx.cell('Valuation', 'J26')), xl_ref(ctx.cell('Valuation', 'J27')))

def calc_Valuation_K26(ctx):
    return xl_sub(xl_add(xl_ref(ctx.cell('Valuation', 'K23')), xl_ref(ctx.cell('Valuation', 'K24'))), xl_ref(ctx.cell('Valuation', 'K25')))

def calc_Valuation_I30(ctx):
    return xl_ref(ctx.cell('Valuation', 'H30'))

def calc_Valuation_J34(ctx):
    return xl_mul(xl_ref(ctx.cell('Valuation', 'J33')), xl_ref(ctx.cell('Valuation', 'J28')))

def calc_Valuation_B65(ctx):
    return xl_npv(xl_ref(ctx.cell('Valuation', 'A65')), ctx.range('Valuation', 'C28:J28'))

def calc_Valuation_B66(ctx):
    return xl_npv(xl_ref(ctx.cell('Valuation', 'A66')), ctx.range('Valuation', 'C28:J28'))

def calc_Valuation_B67(ctx):
    return xl_npv(xl_ref(ctx.cell('Valuation', 'A67')), ctx.range('Valuation', 'C28:J28'))

def calc_Valuation_B68(ctx):
    return xl_npv(xl_ref(ctx.cell('Valuation', 'A68')), ctx.range('Valuation', 'C28:J28'))

def calc_Valuation_B69(ctx):
    return xl_npv(xl_ref(ctx.cell('Valuation', 'A69')), ctx.range('Valuation', 'C28:J28'))

def calc_Valuation_B75(ctx):
    return xl_npv(xl_ref(ctx.cell('Valuation', 'A75')), ctx.range('Valuation', 'C28:J28'))

def calc_Valuation_B76(ctx):
    return xl_npv(xl_ref(ctx.cell('Valuation', 'A76')), ctx.range('Valuation', 'C28:J28'))

def calc_Valuation_B77(ctx):
    return xl_npv(xl_ref(ctx.cell('Valuation', 'A77')), ctx.range('Valuation', 'C28:J28'))

def calc_Valuation_B78(ctx):
    return xl_npv(xl_ref(ctx.cell('Valuation', 'A78')), ctx.range('Valuation', 'C28:J28'))

def calc_Valuation_B79(ctx):
    return xl_npv(xl_ref(ctx.cell('Valuation', 'A79')), ctx.range('Valuation', 'C28:J28'))

def calc_Valuation_K28(ctx):
    return xl_add(xl_ref(ctx.cell('Valuation', 'K26')), xl_ref(ctx.cell('Valuation', 'K27')))

def calc_Valuation_J30(ctx):
    return xl_ref(ctx.cell('Valuation', 'I30'))

def calc_Valuation_B52(ctx):
    return xl_sum(ctx.range('Valuation', 'C34:J34'))

def calc_Valuation_H65(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B65')), xl_ref(ctx.cell('Valuation', 'D65'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_I65(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B65')), xl_ref(ctx.cell('Valuation', 'E65'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_J65(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B65')), xl_ref(ctx.cell('Valuation', 'F65'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_H66(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B66')), xl_ref(ctx.cell('Valuation', 'D66'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_I66(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B66')), xl_ref(ctx.cell('Valuation', 'E66'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_J66(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B66')), xl_ref(ctx.cell('Valuation', 'F66'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_H67(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B67')), xl_ref(ctx.cell('Valuation', 'D67'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_I67(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B67')), xl_ref(ctx.cell('Valuation', 'E67'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_J67(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B67')), xl_ref(ctx.cell('Valuation', 'F67'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_H68(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B68')), xl_ref(ctx.cell('Valuation', 'D68'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_I68(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B68')), xl_ref(ctx.cell('Valuation', 'E68'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_J68(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B68')), xl_ref(ctx.cell('Valuation', 'F68'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_H69(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B69')), xl_ref(ctx.cell('Valuation', 'D69'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_I69(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B69')), xl_ref(ctx.cell('Valuation', 'E69'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_J69(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B69')), xl_ref(ctx.cell('Valuation', 'F69'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_H75(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B75')), xl_ref(ctx.cell('Valuation', 'D75'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_I75(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B75')), xl_ref(ctx.cell('Valuation', 'E75'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_J75(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B75')), xl_ref(ctx.cell('Valuation', 'F75'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_H76(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B76')), xl_ref(ctx.cell('Valuation', 'D76'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_I76(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B76')), xl_ref(ctx.cell('Valuation', 'E76'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_J76(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B76')), xl_ref(ctx.cell('Valuation', 'F76'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_H77(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B77')), xl_ref(ctx.cell('Valuation', 'D77'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_I77(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B77')), xl_ref(ctx.cell('Valuation', 'E77'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_J77(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B77')), xl_ref(ctx.cell('Valuation', 'F77'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_H78(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B78')), xl_ref(ctx.cell('Valuation', 'D78'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_I78(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B78')), xl_ref(ctx.cell('Valuation', 'E78'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_J78(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B78')), xl_ref(ctx.cell('Valuation', 'F78'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_H79(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B79')), xl_ref(ctx.cell('Valuation', 'D79'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_I79(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B79')), xl_ref(ctx.cell('Valuation', 'E79'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_J79(ctx):
    return xl_div(xl_add(xl_add(xl_ref(ctx.cell('Valuation', 'B79')), xl_ref(ctx.cell('Valuation', 'F79'))), xl_ref(ctx.cell('Valuation', 'G55'))), xl_ref(ctx.cell('Valuation', 'K53')))

def calc_Valuation_B55(ctx):
    return xl_mul(xl_div(xl_ref(ctx.cell('Valuation', 'K28')), xl_sub(xl_ref(ctx.cell('Valuation', 'B53')), xl_ref(ctx.cell('Valuation', 'B54')))), xl_ref(ctx.cell('Valuation', 'J33')))

def calc_Valuation_K30(ctx):
    return xl_ref(ctx.cell('Valuation', 'J30'))

def calc_Valuation_G52(ctx):
    return xl_add(xl_ref(ctx.cell('Valuation', 'B55')), xl_ref(ctx.cell('Valuation', 'B52')))

def calc_Valuation_B56(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'B55')), xl_add(xl_ref(ctx.cell('Valuation', 'B55')), xl_ref(ctx.cell('Valuation', 'B52'))))

def calc_Valuation_C64(ctx):
    return xl_ref(ctx.cell('Valuation', 'B55'))

def calc_Valuation_G56(ctx):
    return xl_add(xl_ref(ctx.cell('Valuation', 'G52')), xl_ref(ctx.cell('Valuation', 'G55')))

def calc_Valuation_K52(ctx):
    return xl_ref(ctx.cell('Valuation', 'G56'))

def calc_Valuation_K55(ctx):
    return xl_div(xl_ref(ctx.cell('Valuation', 'K52')), xl_ref(ctx.cell('Valuation', 'K53')))

FORMULA_FUNCS = {
    ('COMPANY OVERVIEW', 'G23'): calc_COMPANY_OVERVIEW_G23,
    ('COMPANY OVERVIEW', 'H23'): calc_COMPANY_OVERVIEW_H23,
    ('COMPANY OVERVIEW', 'I23'): calc_COMPANY_OVERVIEW_I23,
    ('COMPANY OVERVIEW', 'J23'): calc_COMPANY_OVERVIEW_J23,
    ('COMPANY OVERVIEW', 'E28'): calc_COMPANY_OVERVIEW_E28,
    ('COMPANY OVERVIEW', 'G28'): calc_COMPANY_OVERVIEW_G28,
    ('COMPANY OVERVIEW', 'H28'): calc_COMPANY_OVERVIEW_H28,
    ('COMPANY OVERVIEW', 'I28'): calc_COMPANY_OVERVIEW_I28,
    ('COMPANY OVERVIEW', 'J28'): calc_COMPANY_OVERVIEW_J28,
    ('COMPANY OVERVIEW', 'B39'): calc_COMPANY_OVERVIEW_B39,
    ('COMPANY OVERVIEW', 'B57'): calc_COMPANY_OVERVIEW_B57,
    ('COMPANY OVERVIEW', 'C57'): calc_COMPANY_OVERVIEW_C57,
    ('COMPANY OVERVIEW', 'D57'): calc_COMPANY_OVERVIEW_D57,
    ('COMPANY OVERVIEW', 'E57'): calc_COMPANY_OVERVIEW_E57,
    ('COMPANY OVERVIEW', 'F57'): calc_COMPANY_OVERVIEW_F57,
    ('COMPANY OVERVIEW', 'G57'): calc_COMPANY_OVERVIEW_G57,
    ('PRESENTATION', 'F8'): calc_PRESENTATION_F8,
    ('PRESENTATION', 'K8'): calc_PRESENTATION_K8,
    ('PRESENTATION', 'P8'): calc_PRESENTATION_P8,
    ('PRESENTATION', 'T8'): calc_PRESENTATION_T8,
    ('PRESENTATION', 'F9'): calc_PRESENTATION_F9,
    ('PRESENTATION', 'K9'): calc_PRESENTATION_K9,
    ('PRESENTATION', 'P9'): calc_PRESENTATION_P9,
    ('PRESENTATION', 'T9'): calc_PRESENTATION_T9,
    ('PRESENTATION', 'F10'): calc_PRESENTATION_F10,
    ('PRESENTATION', 'K10'): calc_PRESENTATION_K10,
    ('PRESENTATION', 'P10'): calc_PRESENTATION_P10,
    ('PRESENTATION', 'T10'): calc_PRESENTATION_T10,
    ('PRESENTATION', 'F11'): calc_PRESENTATION_F11,
    ('PRESENTATION', 'K11'): calc_PRESENTATION_K11,
    ('PRESENTATION', 'P11'): calc_PRESENTATION_P11,
    ('PRESENTATION', 'T11'): calc_PRESENTATION_T11,
    ('PRESENTATION', 'F12'): calc_PRESENTATION_F12,
    ('PRESENTATION', 'K12'): calc_PRESENTATION_K12,
    ('PRESENTATION', 'P12'): calc_PRESENTATION_P12,
    ('PRESENTATION', 'T12'): calc_PRESENTATION_T12,
    ('PRESENTATION', 'F13'): calc_PRESENTATION_F13,
    ('PRESENTATION', 'K13'): calc_PRESENTATION_K13,
    ('PRESENTATION', 'P13'): calc_PRESENTATION_P13,
    ('PRESENTATION', 'T13'): calc_PRESENTATION_T13,
    ('PRESENTATION', 'F14'): calc_PRESENTATION_F14,
    ('PRESENTATION', 'K14'): calc_PRESENTATION_K14,
    ('PRESENTATION', 'P14'): calc_PRESENTATION_P14,
    ('PRESENTATION', 'T14'): calc_PRESENTATION_T14,
    ('PRESENTATION', 'B15'): calc_PRESENTATION_B15,
    ('PRESENTATION', 'C15'): calc_PRESENTATION_C15,
    ('PRESENTATION', 'D15'): calc_PRESENTATION_D15,
    ('PRESENTATION', 'E15'): calc_PRESENTATION_E15,
    ('PRESENTATION', 'G15'): calc_PRESENTATION_G15,
    ('PRESENTATION', 'H15'): calc_PRESENTATION_H15,
    ('PRESENTATION', 'I15'): calc_PRESENTATION_I15,
    ('PRESENTATION', 'J15'): calc_PRESENTATION_J15,
    ('PRESENTATION', 'L15'): calc_PRESENTATION_L15,
    ('PRESENTATION', 'M15'): calc_PRESENTATION_M15,
    ('PRESENTATION', 'N15'): calc_PRESENTATION_N15,
    ('PRESENTATION', 'O15'): calc_PRESENTATION_O15,
    ('PRESENTATION', 'Q15'): calc_PRESENTATION_Q15,
    ('PRESENTATION', 'R15'): calc_PRESENTATION_R15,
    ('PRESENTATION', 'S15'): calc_PRESENTATION_S15,
    ('PRESENTATION', 'V15'): calc_PRESENTATION_V15,
    ('PRESENTATION', 'W15'): calc_PRESENTATION_W15,
    ('PRESENTATION', 'X15'): calc_PRESENTATION_X15,
    ('PRESENTATION', 'Y15'): calc_PRESENTATION_Y15,
    ('PRESENTATION', 'F16'): calc_PRESENTATION_F16,
    ('PRESENTATION', 'P16'): calc_PRESENTATION_P16,
    ('PRESENTATION', 'U16'): calc_PRESENTATION_U16,
    ('PRESENTATION', 'D25'): calc_PRESENTATION_D25,
    ('PRESENTATION', 'K25'): calc_PRESENTATION_K25,
    ('PRESENTATION', 'R25'): calc_PRESENTATION_R25,
    ('PRESENTATION', 'Y25'): calc_PRESENTATION_Y25,
    ('PRESENTATION', 'B26'): calc_PRESENTATION_B26,
    ('PRESENTATION', 'C26'): calc_PRESENTATION_C26,
    ('PRESENTATION', 'E26'): calc_PRESENTATION_E26,
    ('PRESENTATION', 'G26'): calc_PRESENTATION_G26,
    ('PRESENTATION', 'I26'): calc_PRESENTATION_I26,
    ('PRESENTATION', 'J26'): calc_PRESENTATION_J26,
    ('PRESENTATION', 'L26'): calc_PRESENTATION_L26,
    ('PRESENTATION', 'N26'): calc_PRESENTATION_N26,
    ('PRESENTATION', 'P26'): calc_PRESENTATION_P26,
    ('PRESENTATION', 'Q26'): calc_PRESENTATION_Q26,
    ('PRESENTATION', 'S26'): calc_PRESENTATION_S26,
    ('PRESENTATION', 'U26'): calc_PRESENTATION_U26,
    ('PRESENTATION', 'W26'): calc_PRESENTATION_W26,
    ('PRESENTATION', 'X26'): calc_PRESENTATION_X26,
    ('PRESENTATION', 'Z26'): calc_PRESENTATION_Z26,
    ('PRESENTATION', 'AB26'): calc_PRESENTATION_AB26,
    ('PRESENTATION', 'AD26'): calc_PRESENTATION_AD26,
    ('PRESENTATION', 'AE26'): calc_PRESENTATION_AE26,
    ('PRESENTATION', 'AF26'): calc_PRESENTATION_AF26,
    ('PRESENTATION', 'AG26'): calc_PRESENTATION_AG26,
    ('PRESENTATION', 'D27'): calc_PRESENTATION_D27,
    ('PRESENTATION', 'K27'): calc_PRESENTATION_K27,
    ('PRESENTATION', 'R27'): calc_PRESENTATION_R27,
    ('PRESENTATION', 'Y27'): calc_PRESENTATION_Y27,
    ('PRESENTATION', 'D28'): calc_PRESENTATION_D28,
    ('PRESENTATION', 'K28'): calc_PRESENTATION_K28,
    ('PRESENTATION', 'R28'): calc_PRESENTATION_R28,
    ('PRESENTATION', 'Y28'): calc_PRESENTATION_Y28,
    ('PRESENTATION', 'D29'): calc_PRESENTATION_D29,
    ('PRESENTATION', 'K29'): calc_PRESENTATION_K29,
    ('PRESENTATION', 'R29'): calc_PRESENTATION_R29,
    ('PRESENTATION', 'Y29'): calc_PRESENTATION_Y29,
    ('PRESENTATION', 'D30'): calc_PRESENTATION_D30,
    ('PRESENTATION', 'K30'): calc_PRESENTATION_K30,
    ('PRESENTATION', 'R30'): calc_PRESENTATION_R30,
    ('PRESENTATION', 'Y30'): calc_PRESENTATION_Y30,
    ('PRESENTATION', 'D31'): calc_PRESENTATION_D31,
    ('PRESENTATION', 'K31'): calc_PRESENTATION_K31,
    ('PRESENTATION', 'R31'): calc_PRESENTATION_R31,
    ('PRESENTATION', 'Y31'): calc_PRESENTATION_Y31,
    ('PRESENTATION', 'K32'): calc_PRESENTATION_K32,
    ('PRESENTATION', 'D34'): calc_PRESENTATION_D34,
    ('PRESENTATION', 'K34'): calc_PRESENTATION_K34,
    ('PRESENTATION', 'R34'): calc_PRESENTATION_R34,
    ('PRESENTATION', 'Y34'): calc_PRESENTATION_Y34,
    ('PRESENTATION', 'D35'): calc_PRESENTATION_D35,
    ('PRESENTATION', 'K35'): calc_PRESENTATION_K35,
    ('PRESENTATION', 'R35'): calc_PRESENTATION_R35,
    ('PRESENTATION', 'Y35'): calc_PRESENTATION_Y35,
    ('PRESENTATION', 'D37'): calc_PRESENTATION_D37,
    ('PRESENTATION', 'K37'): calc_PRESENTATION_K37,
    ('PRESENTATION', 'R37'): calc_PRESENTATION_R37,
    ('PRESENTATION', 'V37'): calc_PRESENTATION_V37,
    ('PRESENTATION', 'Y37'): calc_PRESENTATION_Y37,
    ('PRESENTATION', 'M39'): calc_PRESENTATION_M39,
    ('PRESENTATION', 'R39'): calc_PRESENTATION_R39,
    ('PRESENTATION', 'Y39'): calc_PRESENTATION_Y39,
    ('PRESENTATION', 'D41'): calc_PRESENTATION_D41,
    ('PRESENTATION', 'K41'): calc_PRESENTATION_K41,
    ('PRESENTATION', 'R41'): calc_PRESENTATION_R41,
    ('PRESENTATION', 'Y41'): calc_PRESENTATION_Y41,
    ('PRESENTATION', 'D42'): calc_PRESENTATION_D42,
    ('PRESENTATION', 'K42'): calc_PRESENTATION_K42,
    ('PRESENTATION', 'R42'): calc_PRESENTATION_R42,
    ('PRESENTATION', 'Y42'): calc_PRESENTATION_Y42,
    ('PRESENTATION', 'D43'): calc_PRESENTATION_D43,
    ('PRESENTATION', 'K43'): calc_PRESENTATION_K43,
    ('PRESENTATION', 'R43'): calc_PRESENTATION_R43,
    ('PRESENTATION', 'Y43'): calc_PRESENTATION_Y43,
    ('Segment Revenue Model', 'F8'): calc_Segment_Revenue_Model_F8,
    ('Segment Revenue Model', 'K8'): calc_Segment_Revenue_Model_K8,
    ('Segment Revenue Model', 'P8'): calc_Segment_Revenue_Model_P8,
    ('Segment Revenue Model', 'F9'): calc_Segment_Revenue_Model_F9,
    ('Segment Revenue Model', 'K9'): calc_Segment_Revenue_Model_K9,
    ('Segment Revenue Model', 'P9'): calc_Segment_Revenue_Model_P9,
    ('Segment Revenue Model', 'F10'): calc_Segment_Revenue_Model_F10,
    ('Segment Revenue Model', 'K10'): calc_Segment_Revenue_Model_K10,
    ('Segment Revenue Model', 'P10'): calc_Segment_Revenue_Model_P10,
    ('Segment Revenue Model', 'F11'): calc_Segment_Revenue_Model_F11,
    ('Segment Revenue Model', 'K11'): calc_Segment_Revenue_Model_K11,
    ('Segment Revenue Model', 'P11'): calc_Segment_Revenue_Model_P11,
    ('Segment Revenue Model', 'F12'): calc_Segment_Revenue_Model_F12,
    ('Segment Revenue Model', 'K12'): calc_Segment_Revenue_Model_K12,
    ('Segment Revenue Model', 'P12'): calc_Segment_Revenue_Model_P12,
    ('Segment Revenue Model', 'U12'): calc_Segment_Revenue_Model_U12,
    ('Segment Revenue Model', 'F13'): calc_Segment_Revenue_Model_F13,
    ('Segment Revenue Model', 'K13'): calc_Segment_Revenue_Model_K13,
    ('Segment Revenue Model', 'P13'): calc_Segment_Revenue_Model_P13,
    ('Segment Revenue Model', 'F14'): calc_Segment_Revenue_Model_F14,
    ('Segment Revenue Model', 'K14'): calc_Segment_Revenue_Model_K14,
    ('Segment Revenue Model', 'P14'): calc_Segment_Revenue_Model_P14,
    ('Segment Revenue Model', 'B15'): calc_Segment_Revenue_Model_B15,
    ('Segment Revenue Model', 'C15'): calc_Segment_Revenue_Model_C15,
    ('Segment Revenue Model', 'D15'): calc_Segment_Revenue_Model_D15,
    ('Segment Revenue Model', 'E15'): calc_Segment_Revenue_Model_E15,
    ('Segment Revenue Model', 'G15'): calc_Segment_Revenue_Model_G15,
    ('Segment Revenue Model', 'H15'): calc_Segment_Revenue_Model_H15,
    ('Segment Revenue Model', 'I15'): calc_Segment_Revenue_Model_I15,
    ('Segment Revenue Model', 'J15'): calc_Segment_Revenue_Model_J15,
    ('Segment Revenue Model', 'L15'): calc_Segment_Revenue_Model_L15,
    ('Segment Revenue Model', 'M15'): calc_Segment_Revenue_Model_M15,
    ('Segment Revenue Model', 'N15'): calc_Segment_Revenue_Model_N15,
    ('Segment Revenue Model', 'O15'): calc_Segment_Revenue_Model_O15,
    ('Segment Revenue Model', 'Q15'): calc_Segment_Revenue_Model_Q15,
    ('Segment Revenue Model', 'R15'): calc_Segment_Revenue_Model_R15,
    ('Segment Revenue Model', 'S15'): calc_Segment_Revenue_Model_S15,
    ('Segment Revenue Model', 'F16'): calc_Segment_Revenue_Model_F16,
    ('Segment Revenue Model', 'P16'): calc_Segment_Revenue_Model_P16,
    ('Segment Revenue Model', 'U16'): calc_Segment_Revenue_Model_U16,
    ('Segment Revenue Model', 'G24'): calc_Segment_Revenue_Model_G24,
    ('Segment Revenue Model', 'H24'): calc_Segment_Revenue_Model_H24,
    ('Segment Revenue Model', 'I24'): calc_Segment_Revenue_Model_I24,
    ('Segment Revenue Model', 'J24'): calc_Segment_Revenue_Model_J24,
    ('Segment Revenue Model', 'L24'): calc_Segment_Revenue_Model_L24,
    ('Segment Revenue Model', 'M24'): calc_Segment_Revenue_Model_M24,
    ('Segment Revenue Model', 'N24'): calc_Segment_Revenue_Model_N24,
    ('Segment Revenue Model', 'O24'): calc_Segment_Revenue_Model_O24,
    ('Segment Revenue Model', 'Q24'): calc_Segment_Revenue_Model_Q24,
    ('Segment Revenue Model', 'R24'): calc_Segment_Revenue_Model_R24,
    ('Segment Revenue Model', 'S24'): calc_Segment_Revenue_Model_S24,
    ('Segment Revenue Model', 'G25'): calc_Segment_Revenue_Model_G25,
    ('Segment Revenue Model', 'H25'): calc_Segment_Revenue_Model_H25,
    ('Segment Revenue Model', 'I25'): calc_Segment_Revenue_Model_I25,
    ('Segment Revenue Model', 'J25'): calc_Segment_Revenue_Model_J25,
    ('Segment Revenue Model', 'L25'): calc_Segment_Revenue_Model_L25,
    ('Segment Revenue Model', 'M25'): calc_Segment_Revenue_Model_M25,
    ('Segment Revenue Model', 'N25'): calc_Segment_Revenue_Model_N25,
    ('Segment Revenue Model', 'O25'): calc_Segment_Revenue_Model_O25,
    ('Segment Revenue Model', 'Q25'): calc_Segment_Revenue_Model_Q25,
    ('Segment Revenue Model', 'R25'): calc_Segment_Revenue_Model_R25,
    ('Segment Revenue Model', 'S25'): calc_Segment_Revenue_Model_S25,
    ('Segment Revenue Model', 'G26'): calc_Segment_Revenue_Model_G26,
    ('Segment Revenue Model', 'H26'): calc_Segment_Revenue_Model_H26,
    ('Segment Revenue Model', 'I26'): calc_Segment_Revenue_Model_I26,
    ('Segment Revenue Model', 'J26'): calc_Segment_Revenue_Model_J26,
    ('Segment Revenue Model', 'L26'): calc_Segment_Revenue_Model_L26,
    ('Segment Revenue Model', 'M26'): calc_Segment_Revenue_Model_M26,
    ('Segment Revenue Model', 'N26'): calc_Segment_Revenue_Model_N26,
    ('Segment Revenue Model', 'O26'): calc_Segment_Revenue_Model_O26,
    ('Segment Revenue Model', 'Q26'): calc_Segment_Revenue_Model_Q26,
    ('Segment Revenue Model', 'R26'): calc_Segment_Revenue_Model_R26,
    ('Segment Revenue Model', 'S26'): calc_Segment_Revenue_Model_S26,
    ('Segment Revenue Model', 'G27'): calc_Segment_Revenue_Model_G27,
    ('Segment Revenue Model', 'H27'): calc_Segment_Revenue_Model_H27,
    ('Segment Revenue Model', 'I27'): calc_Segment_Revenue_Model_I27,
    ('Segment Revenue Model', 'J27'): calc_Segment_Revenue_Model_J27,
    ('Segment Revenue Model', 'L27'): calc_Segment_Revenue_Model_L27,
    ('Segment Revenue Model', 'M27'): calc_Segment_Revenue_Model_M27,
    ('Segment Revenue Model', 'N27'): calc_Segment_Revenue_Model_N27,
    ('Segment Revenue Model', 'O27'): calc_Segment_Revenue_Model_O27,
    ('Segment Revenue Model', 'Q27'): calc_Segment_Revenue_Model_Q27,
    ('Segment Revenue Model', 'R27'): calc_Segment_Revenue_Model_R27,
    ('Segment Revenue Model', 'S27'): calc_Segment_Revenue_Model_S27,
    ('Segment Revenue Model', 'G28'): calc_Segment_Revenue_Model_G28,
    ('Segment Revenue Model', 'H28'): calc_Segment_Revenue_Model_H28,
    ('Segment Revenue Model', 'I28'): calc_Segment_Revenue_Model_I28,
    ('Segment Revenue Model', 'J28'): calc_Segment_Revenue_Model_J28,
    ('Segment Revenue Model', 'L28'): calc_Segment_Revenue_Model_L28,
    ('Segment Revenue Model', 'M28'): calc_Segment_Revenue_Model_M28,
    ('Segment Revenue Model', 'N28'): calc_Segment_Revenue_Model_N28,
    ('Segment Revenue Model', 'O28'): calc_Segment_Revenue_Model_O28,
    ('Segment Revenue Model', 'Q28'): calc_Segment_Revenue_Model_Q28,
    ('Segment Revenue Model', 'R28'): calc_Segment_Revenue_Model_R28,
    ('Segment Revenue Model', 'S28'): calc_Segment_Revenue_Model_S28,
    ('Segment Revenue Model', 'G29'): calc_Segment_Revenue_Model_G29,
    ('Segment Revenue Model', 'H29'): calc_Segment_Revenue_Model_H29,
    ('Segment Revenue Model', 'I29'): calc_Segment_Revenue_Model_I29,
    ('Segment Revenue Model', 'J29'): calc_Segment_Revenue_Model_J29,
    ('Segment Revenue Model', 'L29'): calc_Segment_Revenue_Model_L29,
    ('Segment Revenue Model', 'M29'): calc_Segment_Revenue_Model_M29,
    ('Segment Revenue Model', 'N29'): calc_Segment_Revenue_Model_N29,
    ('Segment Revenue Model', 'O29'): calc_Segment_Revenue_Model_O29,
    ('Segment Revenue Model', 'Q29'): calc_Segment_Revenue_Model_Q29,
    ('Segment Revenue Model', 'R29'): calc_Segment_Revenue_Model_R29,
    ('Segment Revenue Model', 'S29'): calc_Segment_Revenue_Model_S29,
    ('Segment Revenue Model', 'G30'): calc_Segment_Revenue_Model_G30,
    ('Segment Revenue Model', 'H30'): calc_Segment_Revenue_Model_H30,
    ('Segment Revenue Model', 'I30'): calc_Segment_Revenue_Model_I30,
    ('Segment Revenue Model', 'J30'): calc_Segment_Revenue_Model_J30,
    ('Segment Revenue Model', 'L30'): calc_Segment_Revenue_Model_L30,
    ('Segment Revenue Model', 'M30'): calc_Segment_Revenue_Model_M30,
    ('Segment Revenue Model', 'N30'): calc_Segment_Revenue_Model_N30,
    ('Segment Revenue Model', 'O30'): calc_Segment_Revenue_Model_O30,
    ('Segment Revenue Model', 'Q30'): calc_Segment_Revenue_Model_Q30,
    ('Segment Revenue Model', 'R30'): calc_Segment_Revenue_Model_R30,
    ('Segment Revenue Model', 'S30'): calc_Segment_Revenue_Model_S30,
    ('Segment Revenue Model', 'G32'): calc_Segment_Revenue_Model_G32,
    ('Segment Revenue Model', 'H32'): calc_Segment_Revenue_Model_H32,
    ('Segment Revenue Model', 'I32'): calc_Segment_Revenue_Model_I32,
    ('Segment Revenue Model', 'J32'): calc_Segment_Revenue_Model_J32,
    ('Segment Revenue Model', 'L32'): calc_Segment_Revenue_Model_L32,
    ('Segment Revenue Model', 'M32'): calc_Segment_Revenue_Model_M32,
    ('Segment Revenue Model', 'N32'): calc_Segment_Revenue_Model_N32,
    ('Segment Revenue Model', 'O32'): calc_Segment_Revenue_Model_O32,
    ('Segment Revenue Model', 'Q32'): calc_Segment_Revenue_Model_Q32,
    ('Segment Revenue Model', 'R32'): calc_Segment_Revenue_Model_R32,
    ('Segment Revenue Model', 'S32'): calc_Segment_Revenue_Model_S32,
    ('Segment Revenue Model', 'T32'): calc_Segment_Revenue_Model_T32,
    ('Segment Revenue Model', 'F51'): calc_Segment_Revenue_Model_F51,
    ('Segment Revenue Model', 'K51'): calc_Segment_Revenue_Model_K51,
    ('Segment Revenue Model', 'P51'): calc_Segment_Revenue_Model_P51,
    ('Segment Revenue Model', 'U51'): calc_Segment_Revenue_Model_U51,
    ('Segment Revenue Model', 'F52'): calc_Segment_Revenue_Model_F52,
    ('Segment Revenue Model', 'P52'): calc_Segment_Revenue_Model_P52,
    ('Segment Revenue Model', 'U52'): calc_Segment_Revenue_Model_U52,
    ('Segment Revenue Model', 'F53'): calc_Segment_Revenue_Model_F53,
    ('Segment Revenue Model', 'K53'): calc_Segment_Revenue_Model_K53,
    ('Segment Revenue Model', 'P53'): calc_Segment_Revenue_Model_P53,
    ('Segment Revenue Model', 'U53'): calc_Segment_Revenue_Model_U53,
    ('Segment Revenue Model', 'F54'): calc_Segment_Revenue_Model_F54,
    ('Segment Revenue Model', 'K54'): calc_Segment_Revenue_Model_K54,
    ('Segment Revenue Model', 'P54'): calc_Segment_Revenue_Model_P54,
    ('Segment Revenue Model', 'U54'): calc_Segment_Revenue_Model_U54,
    ('Segment Revenue Model', 'F55'): calc_Segment_Revenue_Model_F55,
    ('Segment Revenue Model', 'K55'): calc_Segment_Revenue_Model_K55,
    ('Segment Revenue Model', 'P55'): calc_Segment_Revenue_Model_P55,
    ('Segment Revenue Model', 'U55'): calc_Segment_Revenue_Model_U55,
    ('Segment Revenue Model', 'F56'): calc_Segment_Revenue_Model_F56,
    ('Segment Revenue Model', 'K56'): calc_Segment_Revenue_Model_K56,
    ('Segment Revenue Model', 'P56'): calc_Segment_Revenue_Model_P56,
    ('Segment Revenue Model', 'U56'): calc_Segment_Revenue_Model_U56,
    ('Segment Revenue Model', 'F57'): calc_Segment_Revenue_Model_F57,
    ('Segment Revenue Model', 'K57'): calc_Segment_Revenue_Model_K57,
    ('Segment Revenue Model', 'P57'): calc_Segment_Revenue_Model_P57,
    ('Segment Revenue Model', 'U57'): calc_Segment_Revenue_Model_U57,
    ('Segment Revenue Model', 'B58'): calc_Segment_Revenue_Model_B58,
    ('Segment Revenue Model', 'C58'): calc_Segment_Revenue_Model_C58,
    ('Segment Revenue Model', 'D58'): calc_Segment_Revenue_Model_D58,
    ('Segment Revenue Model', 'E58'): calc_Segment_Revenue_Model_E58,
    ('Segment Revenue Model', 'G58'): calc_Segment_Revenue_Model_G58,
    ('Segment Revenue Model', 'H58'): calc_Segment_Revenue_Model_H58,
    ('Segment Revenue Model', 'I58'): calc_Segment_Revenue_Model_I58,
    ('Segment Revenue Model', 'J58'): calc_Segment_Revenue_Model_J58,
    ('Segment Revenue Model', 'L58'): calc_Segment_Revenue_Model_L58,
    ('Segment Revenue Model', 'M58'): calc_Segment_Revenue_Model_M58,
    ('Segment Revenue Model', 'N58'): calc_Segment_Revenue_Model_N58,
    ('Segment Revenue Model', 'O58'): calc_Segment_Revenue_Model_O58,
    ('Segment Revenue Model', 'Q58'): calc_Segment_Revenue_Model_Q58,
    ('Segment Revenue Model', 'R58'): calc_Segment_Revenue_Model_R58,
    ('Segment Revenue Model', 'S58'): calc_Segment_Revenue_Model_S58,
    ('Segment Revenue Model', 'T58'): calc_Segment_Revenue_Model_T58,
    ('Segment Revenue Model', 'V58'): calc_Segment_Revenue_Model_V58,
    ('Segment Revenue Model', 'W58'): calc_Segment_Revenue_Model_W58,
    ('Segment Revenue Model', 'X58'): calc_Segment_Revenue_Model_X58,
    ('Segment Revenue Model', 'Y58'): calc_Segment_Revenue_Model_Y58,
    ('Segment Revenue Model', 'F59'): calc_Segment_Revenue_Model_F59,
    ('Segment Revenue Model', 'K59'): calc_Segment_Revenue_Model_K59,
    ('Segment Revenue Model', 'P59'): calc_Segment_Revenue_Model_P59,
    ('Segment Revenue Model', 'U59'): calc_Segment_Revenue_Model_U59,
    ('Segment Revenue Model', 'F60'): calc_Segment_Revenue_Model_F60,
    ('Segment Revenue Model', 'K60'): calc_Segment_Revenue_Model_K60,
    ('Segment Revenue Model', 'P60'): calc_Segment_Revenue_Model_P60,
    ('Segment Revenue Model', 'U60'): calc_Segment_Revenue_Model_U60,
    ('Segment Revenue Model', 'L61'): calc_Segment_Revenue_Model_L61,
    ('Segment Revenue Model', 'F65'): calc_Segment_Revenue_Model_F65,
    ('Segment Revenue Model', 'K65'): calc_Segment_Revenue_Model_K65,
    ('Segment Revenue Model', 'P65'): calc_Segment_Revenue_Model_P65,
    ('Segment Revenue Model', 'U65'): calc_Segment_Revenue_Model_U65,
    ('Segment Revenue Model', 'F66'): calc_Segment_Revenue_Model_F66,
    ('Segment Revenue Model', 'K66'): calc_Segment_Revenue_Model_K66,
    ('Segment Revenue Model', 'P66'): calc_Segment_Revenue_Model_P66,
    ('Segment Revenue Model', 'U66'): calc_Segment_Revenue_Model_U66,
    ('Segment Revenue Model', 'F67'): calc_Segment_Revenue_Model_F67,
    ('Segment Revenue Model', 'K67'): calc_Segment_Revenue_Model_K67,
    ('Segment Revenue Model', 'P67'): calc_Segment_Revenue_Model_P67,
    ('Segment Revenue Model', 'U67'): calc_Segment_Revenue_Model_U67,
    ('Segment Revenue Model', 'F68'): calc_Segment_Revenue_Model_F68,
    ('Segment Revenue Model', 'K68'): calc_Segment_Revenue_Model_K68,
    ('Segment Revenue Model', 'P68'): calc_Segment_Revenue_Model_P68,
    ('Segment Revenue Model', 'U68'): calc_Segment_Revenue_Model_U68,
    ('Segment Revenue Model', 'F69'): calc_Segment_Revenue_Model_F69,
    ('Segment Revenue Model', 'K69'): calc_Segment_Revenue_Model_K69,
    ('Segment Revenue Model', 'P69'): calc_Segment_Revenue_Model_P69,
    ('Segment Revenue Model', 'U69'): calc_Segment_Revenue_Model_U69,
    ('Segment Revenue Model', 'F70'): calc_Segment_Revenue_Model_F70,
    ('Segment Revenue Model', 'K70'): calc_Segment_Revenue_Model_K70,
    ('Segment Revenue Model', 'P70'): calc_Segment_Revenue_Model_P70,
    ('Segment Revenue Model', 'U70'): calc_Segment_Revenue_Model_U70,
    ('Segment Revenue Model', 'F71'): calc_Segment_Revenue_Model_F71,
    ('Segment Revenue Model', 'K71'): calc_Segment_Revenue_Model_K71,
    ('Segment Revenue Model', 'P71'): calc_Segment_Revenue_Model_P71,
    ('Segment Revenue Model', 'U71'): calc_Segment_Revenue_Model_U71,
    ('Segment Revenue Model', 'B72'): calc_Segment_Revenue_Model_B72,
    ('Segment Revenue Model', 'C72'): calc_Segment_Revenue_Model_C72,
    ('Segment Revenue Model', 'D72'): calc_Segment_Revenue_Model_D72,
    ('Segment Revenue Model', 'E72'): calc_Segment_Revenue_Model_E72,
    ('Segment Revenue Model', 'G72'): calc_Segment_Revenue_Model_G72,
    ('Segment Revenue Model', 'H72'): calc_Segment_Revenue_Model_H72,
    ('Segment Revenue Model', 'I72'): calc_Segment_Revenue_Model_I72,
    ('Segment Revenue Model', 'J72'): calc_Segment_Revenue_Model_J72,
    ('Segment Revenue Model', 'L72'): calc_Segment_Revenue_Model_L72,
    ('Segment Revenue Model', 'M72'): calc_Segment_Revenue_Model_M72,
    ('Segment Revenue Model', 'N72'): calc_Segment_Revenue_Model_N72,
    ('Segment Revenue Model', 'O72'): calc_Segment_Revenue_Model_O72,
    ('Segment Revenue Model', 'Q72'): calc_Segment_Revenue_Model_Q72,
    ('Segment Revenue Model', 'R72'): calc_Segment_Revenue_Model_R72,
    ('Segment Revenue Model', 'S72'): calc_Segment_Revenue_Model_S72,
    ('Segment Revenue Model', 'T72'): calc_Segment_Revenue_Model_T72,
    ('Segment Revenue Model', 'F73'): calc_Segment_Revenue_Model_F73,
    ('Segment Revenue Model', 'K73'): calc_Segment_Revenue_Model_K73,
    ('Segment Revenue Model', 'P73'): calc_Segment_Revenue_Model_P73,
    ('Segment Revenue Model', 'U73'): calc_Segment_Revenue_Model_U73,
    ('INCOME STATEMENT', 'D6'): calc_INCOME_STATEMENT_D6,
    ('INCOME STATEMENT', 'K6'): calc_INCOME_STATEMENT_K6,
    ('INCOME STATEMENT', 'R6'): calc_INCOME_STATEMENT_R6,
    ('INCOME STATEMENT', 'Y6'): calc_INCOME_STATEMENT_Y6,
    ('INCOME STATEMENT', 'B7'): calc_INCOME_STATEMENT_B7,
    ('INCOME STATEMENT', 'C7'): calc_INCOME_STATEMENT_C7,
    ('INCOME STATEMENT', 'E7'): calc_INCOME_STATEMENT_E7,
    ('INCOME STATEMENT', 'G7'): calc_INCOME_STATEMENT_G7,
    ('INCOME STATEMENT', 'P7'): calc_INCOME_STATEMENT_P7,
    ('INCOME STATEMENT', 'Q7'): calc_INCOME_STATEMENT_Q7,
    ('INCOME STATEMENT', 'S7'): calc_INCOME_STATEMENT_S7,
    ('INCOME STATEMENT', 'U7'): calc_INCOME_STATEMENT_U7,
    ('INCOME STATEMENT', 'H8'): calc_INCOME_STATEMENT_H8,
    ('INCOME STATEMENT', 'I8'): calc_INCOME_STATEMENT_I8,
    ('INCOME STATEMENT', 'J8'): calc_INCOME_STATEMENT_J8,
    ('INCOME STATEMENT', 'L8'): calc_INCOME_STATEMENT_L8,
    ('INCOME STATEMENT', 'N8'): calc_INCOME_STATEMENT_N8,
    ('INCOME STATEMENT', 'P8'): calc_INCOME_STATEMENT_P8,
    ('INCOME STATEMENT', 'Q8'): calc_INCOME_STATEMENT_Q8,
    ('INCOME STATEMENT', 'S8'): calc_INCOME_STATEMENT_S8,
    ('INCOME STATEMENT', 'U8'): calc_INCOME_STATEMENT_U8,
    ('INCOME STATEMENT', 'W8'): calc_INCOME_STATEMENT_W8,
    ('INCOME STATEMENT', 'X8'): calc_INCOME_STATEMENT_X8,
    ('INCOME STATEMENT', 'Z8'): calc_INCOME_STATEMENT_Z8,
    ('INCOME STATEMENT', 'AB8'): calc_INCOME_STATEMENT_AB8,
    ('INCOME STATEMENT', 'AE8'): calc_INCOME_STATEMENT_AE8,
    ('INCOME STATEMENT', 'AF8'): calc_INCOME_STATEMENT_AF8,
    ('INCOME STATEMENT', 'AG8'): calc_INCOME_STATEMENT_AG8,
    ('INCOME STATEMENT', 'C9'): calc_INCOME_STATEMENT_C9,
    ('INCOME STATEMENT', 'E9'): calc_INCOME_STATEMENT_E9,
    ('INCOME STATEMENT', 'G9'): calc_INCOME_STATEMENT_G9,
    ('INCOME STATEMENT', 'I9'): calc_INCOME_STATEMENT_I9,
    ('INCOME STATEMENT', 'J9'): calc_INCOME_STATEMENT_J9,
    ('INCOME STATEMENT', 'L9'): calc_INCOME_STATEMENT_L9,
    ('INCOME STATEMENT', 'N9'): calc_INCOME_STATEMENT_N9,
    ('INCOME STATEMENT', 'P9'): calc_INCOME_STATEMENT_P9,
    ('INCOME STATEMENT', 'Q9'): calc_INCOME_STATEMENT_Q9,
    ('INCOME STATEMENT', 'S9'): calc_INCOME_STATEMENT_S9,
    ('INCOME STATEMENT', 'U9'): calc_INCOME_STATEMENT_U9,
    ('INCOME STATEMENT', 'W9'): calc_INCOME_STATEMENT_W9,
    ('INCOME STATEMENT', 'X9'): calc_INCOME_STATEMENT_X9,
    ('INCOME STATEMENT', 'Z9'): calc_INCOME_STATEMENT_Z9,
    ('INCOME STATEMENT', 'AB9'): calc_INCOME_STATEMENT_AB9,
    ('INCOME STATEMENT', 'B10'): calc_INCOME_STATEMENT_B10,
    ('INCOME STATEMENT', 'C10'): calc_INCOME_STATEMENT_C10,
    ('INCOME STATEMENT', 'E10'): calc_INCOME_STATEMENT_E10,
    ('INCOME STATEMENT', 'G10'): calc_INCOME_STATEMENT_G10,
    ('INCOME STATEMENT', 'I10'): calc_INCOME_STATEMENT_I10,
    ('INCOME STATEMENT', 'J10'): calc_INCOME_STATEMENT_J10,
    ('INCOME STATEMENT', 'L10'): calc_INCOME_STATEMENT_L10,
    ('INCOME STATEMENT', 'N10'): calc_INCOME_STATEMENT_N10,
    ('INCOME STATEMENT', 'P10'): calc_INCOME_STATEMENT_P10,
    ('INCOME STATEMENT', 'Q10'): calc_INCOME_STATEMENT_Q10,
    ('INCOME STATEMENT', 'S10'): calc_INCOME_STATEMENT_S10,
    ('INCOME STATEMENT', 'U10'): calc_INCOME_STATEMENT_U10,
    ('INCOME STATEMENT', 'W10'): calc_INCOME_STATEMENT_W10,
    ('INCOME STATEMENT', 'X10'): calc_INCOME_STATEMENT_X10,
    ('INCOME STATEMENT', 'Z10'): calc_INCOME_STATEMENT_Z10,
    ('INCOME STATEMENT', 'AB10'): calc_INCOME_STATEMENT_AB10,
    ('INCOME STATEMENT', 'D11'): calc_INCOME_STATEMENT_D11,
    ('INCOME STATEMENT', 'K11'): calc_INCOME_STATEMENT_K11,
    ('INCOME STATEMENT', 'R11'): calc_INCOME_STATEMENT_R11,
    ('INCOME STATEMENT', 'Y11'): calc_INCOME_STATEMENT_Y11,
    ('INCOME STATEMENT', 'AD11'): calc_INCOME_STATEMENT_AD11,
    ('INCOME STATEMENT', 'AE11'): calc_INCOME_STATEMENT_AE11,
    ('INCOME STATEMENT', 'AF11'): calc_INCOME_STATEMENT_AF11,
    ('INCOME STATEMENT', 'AG11'): calc_INCOME_STATEMENT_AG11,
    ('INCOME STATEMENT', 'B12'): calc_INCOME_STATEMENT_B12,
    ('INCOME STATEMENT', 'C12'): calc_INCOME_STATEMENT_C12,
    ('INCOME STATEMENT', 'E12'): calc_INCOME_STATEMENT_E12,
    ('INCOME STATEMENT', 'G12'): calc_INCOME_STATEMENT_G12,
    ('INCOME STATEMENT', 'H12'): calc_INCOME_STATEMENT_H12,
    ('INCOME STATEMENT', 'I12'): calc_INCOME_STATEMENT_I12,
    ('INCOME STATEMENT', 'J12'): calc_INCOME_STATEMENT_J12,
    ('INCOME STATEMENT', 'L12'): calc_INCOME_STATEMENT_L12,
    ('INCOME STATEMENT', 'N12'): calc_INCOME_STATEMENT_N12,
    ('INCOME STATEMENT', 'P12'): calc_INCOME_STATEMENT_P12,
    ('INCOME STATEMENT', 'Q12'): calc_INCOME_STATEMENT_Q12,
    ('INCOME STATEMENT', 'S12'): calc_INCOME_STATEMENT_S12,
    ('INCOME STATEMENT', 'U12'): calc_INCOME_STATEMENT_U12,
    ('INCOME STATEMENT', 'V12'): calc_INCOME_STATEMENT_V12,
    ('INCOME STATEMENT', 'W12'): calc_INCOME_STATEMENT_W12,
    ('INCOME STATEMENT', 'X12'): calc_INCOME_STATEMENT_X12,
    ('INCOME STATEMENT', 'Z12'): calc_INCOME_STATEMENT_Z12,
    ('INCOME STATEMENT', 'AB12'): calc_INCOME_STATEMENT_AB12,
    ('INCOME STATEMENT', 'D13'): calc_INCOME_STATEMENT_D13,
    ('INCOME STATEMENT', 'K13'): calc_INCOME_STATEMENT_K13,
    ('INCOME STATEMENT', 'R13'): calc_INCOME_STATEMENT_R13,
    ('INCOME STATEMENT', 'Y13'): calc_INCOME_STATEMENT_Y13,
    ('INCOME STATEMENT', 'AD13'): calc_INCOME_STATEMENT_AD13,
    ('INCOME STATEMENT', 'AE13'): calc_INCOME_STATEMENT_AE13,
    ('INCOME STATEMENT', 'AF13'): calc_INCOME_STATEMENT_AF13,
    ('INCOME STATEMENT', 'AG13'): calc_INCOME_STATEMENT_AG13,
    ('INCOME STATEMENT', 'B14'): calc_INCOME_STATEMENT_B14,
    ('INCOME STATEMENT', 'C14'): calc_INCOME_STATEMENT_C14,
    ('INCOME STATEMENT', 'E14'): calc_INCOME_STATEMENT_E14,
    ('INCOME STATEMENT', 'G14'): calc_INCOME_STATEMENT_G14,
    ('INCOME STATEMENT', 'H14'): calc_INCOME_STATEMENT_H14,
    ('INCOME STATEMENT', 'I14'): calc_INCOME_STATEMENT_I14,
    ('INCOME STATEMENT', 'J14'): calc_INCOME_STATEMENT_J14,
    ('INCOME STATEMENT', 'L14'): calc_INCOME_STATEMENT_L14,
    ('INCOME STATEMENT', 'N14'): calc_INCOME_STATEMENT_N14,
    ('INCOME STATEMENT', 'P14'): calc_INCOME_STATEMENT_P14,
    ('INCOME STATEMENT', 'Q14'): calc_INCOME_STATEMENT_Q14,
    ('INCOME STATEMENT', 'S14'): calc_INCOME_STATEMENT_S14,
    ('INCOME STATEMENT', 'U14'): calc_INCOME_STATEMENT_U14,
    ('INCOME STATEMENT', 'V14'): calc_INCOME_STATEMENT_V14,
    ('INCOME STATEMENT', 'W14'): calc_INCOME_STATEMENT_W14,
    ('INCOME STATEMENT', 'X14'): calc_INCOME_STATEMENT_X14,
    ('INCOME STATEMENT', 'Z14'): calc_INCOME_STATEMENT_Z14,
    ('INCOME STATEMENT', 'AB14'): calc_INCOME_STATEMENT_AB14,
    ('INCOME STATEMENT', 'D15'): calc_INCOME_STATEMENT_D15,
    ('INCOME STATEMENT', 'K15'): calc_INCOME_STATEMENT_K15,
    ('INCOME STATEMENT', 'R15'): calc_INCOME_STATEMENT_R15,
    ('INCOME STATEMENT', 'Y15'): calc_INCOME_STATEMENT_Y15,
    ('INCOME STATEMENT', 'AD15'): calc_INCOME_STATEMENT_AD15,
    ('INCOME STATEMENT', 'AE15'): calc_INCOME_STATEMENT_AE15,
    ('INCOME STATEMENT', 'AF15'): calc_INCOME_STATEMENT_AF15,
    ('INCOME STATEMENT', 'AG15'): calc_INCOME_STATEMENT_AG15,
    ('INCOME STATEMENT', 'B16'): calc_INCOME_STATEMENT_B16,
    ('INCOME STATEMENT', 'C16'): calc_INCOME_STATEMENT_C16,
    ('INCOME STATEMENT', 'E16'): calc_INCOME_STATEMENT_E16,
    ('INCOME STATEMENT', 'G16'): calc_INCOME_STATEMENT_G16,
    ('INCOME STATEMENT', 'I16'): calc_INCOME_STATEMENT_I16,
    ('INCOME STATEMENT', 'J16'): calc_INCOME_STATEMENT_J16,
    ('INCOME STATEMENT', 'L16'): calc_INCOME_STATEMENT_L16,
    ('INCOME STATEMENT', 'N16'): calc_INCOME_STATEMENT_N16,
    ('INCOME STATEMENT', 'P16'): calc_INCOME_STATEMENT_P16,
    ('INCOME STATEMENT', 'Q16'): calc_INCOME_STATEMENT_Q16,
    ('INCOME STATEMENT', 'S16'): calc_INCOME_STATEMENT_S16,
    ('INCOME STATEMENT', 'U16'): calc_INCOME_STATEMENT_U16,
    ('INCOME STATEMENT', 'W16'): calc_INCOME_STATEMENT_W16,
    ('INCOME STATEMENT', 'X16'): calc_INCOME_STATEMENT_X16,
    ('INCOME STATEMENT', 'Z16'): calc_INCOME_STATEMENT_Z16,
    ('INCOME STATEMENT', 'AB16'): calc_INCOME_STATEMENT_AB16,
    ('INCOME STATEMENT', 'D17'): calc_INCOME_STATEMENT_D17,
    ('INCOME STATEMENT', 'K17'): calc_INCOME_STATEMENT_K17,
    ('INCOME STATEMENT', 'R17'): calc_INCOME_STATEMENT_R17,
    ('INCOME STATEMENT', 'Y17'): calc_INCOME_STATEMENT_Y17,
    ('INCOME STATEMENT', 'AD17'): calc_INCOME_STATEMENT_AD17,
    ('INCOME STATEMENT', 'AE17'): calc_INCOME_STATEMENT_AE17,
    ('INCOME STATEMENT', 'AF17'): calc_INCOME_STATEMENT_AF17,
    ('INCOME STATEMENT', 'AG17'): calc_INCOME_STATEMENT_AG17,
    ('INCOME STATEMENT', 'B18'): calc_INCOME_STATEMENT_B18,
    ('INCOME STATEMENT', 'C18'): calc_INCOME_STATEMENT_C18,
    ('INCOME STATEMENT', 'E18'): calc_INCOME_STATEMENT_E18,
    ('INCOME STATEMENT', 'G18'): calc_INCOME_STATEMENT_G18,
    ('INCOME STATEMENT', 'H18'): calc_INCOME_STATEMENT_H18,
    ('INCOME STATEMENT', 'I18'): calc_INCOME_STATEMENT_I18,
    ('INCOME STATEMENT', 'J18'): calc_INCOME_STATEMENT_J18,
    ('INCOME STATEMENT', 'L18'): calc_INCOME_STATEMENT_L18,
    ('INCOME STATEMENT', 'N18'): calc_INCOME_STATEMENT_N18,
    ('INCOME STATEMENT', 'P18'): calc_INCOME_STATEMENT_P18,
    ('INCOME STATEMENT', 'Q18'): calc_INCOME_STATEMENT_Q18,
    ('INCOME STATEMENT', 'S18'): calc_INCOME_STATEMENT_S18,
    ('INCOME STATEMENT', 'U18'): calc_INCOME_STATEMENT_U18,
    ('INCOME STATEMENT', 'V18'): calc_INCOME_STATEMENT_V18,
    ('INCOME STATEMENT', 'W18'): calc_INCOME_STATEMENT_W18,
    ('INCOME STATEMENT', 'X18'): calc_INCOME_STATEMENT_X18,
    ('INCOME STATEMENT', 'Z18'): calc_INCOME_STATEMENT_Z18,
    ('INCOME STATEMENT', 'AB18'): calc_INCOME_STATEMENT_AB18,
    ('INCOME STATEMENT', 'D19'): calc_INCOME_STATEMENT_D19,
    ('INCOME STATEMENT', 'K19'): calc_INCOME_STATEMENT_K19,
    ('INCOME STATEMENT', 'R19'): calc_INCOME_STATEMENT_R19,
    ('INCOME STATEMENT', 'Y19'): calc_INCOME_STATEMENT_Y19,
    ('INCOME STATEMENT', 'AD19'): calc_INCOME_STATEMENT_AD19,
    ('INCOME STATEMENT', 'AE19'): calc_INCOME_STATEMENT_AE19,
    ('INCOME STATEMENT', 'AF19'): calc_INCOME_STATEMENT_AF19,
    ('INCOME STATEMENT', 'AG19'): calc_INCOME_STATEMENT_AG19,
    ('INCOME STATEMENT', 'B20'): calc_INCOME_STATEMENT_B20,
    ('INCOME STATEMENT', 'C20'): calc_INCOME_STATEMENT_C20,
    ('INCOME STATEMENT', 'E20'): calc_INCOME_STATEMENT_E20,
    ('INCOME STATEMENT', 'G20'): calc_INCOME_STATEMENT_G20,
    ('INCOME STATEMENT', 'I20'): calc_INCOME_STATEMENT_I20,
    ('INCOME STATEMENT', 'J20'): calc_INCOME_STATEMENT_J20,
    ('INCOME STATEMENT', 'L20'): calc_INCOME_STATEMENT_L20,
    ('INCOME STATEMENT', 'N20'): calc_INCOME_STATEMENT_N20,
    ('INCOME STATEMENT', 'P20'): calc_INCOME_STATEMENT_P20,
    ('INCOME STATEMENT', 'Q20'): calc_INCOME_STATEMENT_Q20,
    ('INCOME STATEMENT', 'S20'): calc_INCOME_STATEMENT_S20,
    ('INCOME STATEMENT', 'U20'): calc_INCOME_STATEMENT_U20,
    ('INCOME STATEMENT', 'W20'): calc_INCOME_STATEMENT_W20,
    ('INCOME STATEMENT', 'X20'): calc_INCOME_STATEMENT_X20,
    ('INCOME STATEMENT', 'Z20'): calc_INCOME_STATEMENT_Z20,
    ('INCOME STATEMENT', 'AB20'): calc_INCOME_STATEMENT_AB20,
    ('INCOME STATEMENT', 'K21'): calc_INCOME_STATEMENT_K21,
    ('INCOME STATEMENT', 'D26'): calc_INCOME_STATEMENT_D26,
    ('INCOME STATEMENT', 'K26'): calc_INCOME_STATEMENT_K26,
    ('INCOME STATEMENT', 'R26'): calc_INCOME_STATEMENT_R26,
    ('INCOME STATEMENT', 'Y26'): calc_INCOME_STATEMENT_Y26,
    ('INCOME STATEMENT', 'D27'): calc_INCOME_STATEMENT_D27,
    ('INCOME STATEMENT', 'K27'): calc_INCOME_STATEMENT_K27,
    ('INCOME STATEMENT', 'R27'): calc_INCOME_STATEMENT_R27,
    ('INCOME STATEMENT', 'Y27'): calc_INCOME_STATEMENT_Y27,
    ('INCOME STATEMENT', 'D29'): calc_INCOME_STATEMENT_D29,
    ('INCOME STATEMENT', 'K29'): calc_INCOME_STATEMENT_K29,
    ('INCOME STATEMENT', 'R29'): calc_INCOME_STATEMENT_R29,
    ('INCOME STATEMENT', 'V29'): calc_INCOME_STATEMENT_V29,
    ('INCOME STATEMENT', 'Y29'): calc_INCOME_STATEMENT_Y29,
    ('INCOME STATEMENT', 'M31'): calc_INCOME_STATEMENT_M31,
    ('INCOME STATEMENT', 'R31'): calc_INCOME_STATEMENT_R31,
    ('INCOME STATEMENT', 'Y31'): calc_INCOME_STATEMENT_Y31,
    ('INCOME STATEMENT', 'D33'): calc_INCOME_STATEMENT_D33,
    ('INCOME STATEMENT', 'K33'): calc_INCOME_STATEMENT_K33,
    ('INCOME STATEMENT', 'R33'): calc_INCOME_STATEMENT_R33,
    ('INCOME STATEMENT', 'Y33'): calc_INCOME_STATEMENT_Y33,
    ('INCOME STATEMENT', 'D34'): calc_INCOME_STATEMENT_D34,
    ('INCOME STATEMENT', 'K34'): calc_INCOME_STATEMENT_K34,
    ('INCOME STATEMENT', 'R34'): calc_INCOME_STATEMENT_R34,
    ('INCOME STATEMENT', 'Y34'): calc_INCOME_STATEMENT_Y34,
    ('INCOME STATEMENT', 'D35'): calc_INCOME_STATEMENT_D35,
    ('INCOME STATEMENT', 'K35'): calc_INCOME_STATEMENT_K35,
    ('INCOME STATEMENT', 'R35'): calc_INCOME_STATEMENT_R35,
    ('INCOME STATEMENT', 'Y35'): calc_INCOME_STATEMENT_Y35,
    ('INCOME STATEMENT', 'D39'): calc_INCOME_STATEMENT_D39,
    ('INCOME STATEMENT', 'H39'): calc_INCOME_STATEMENT_H39,
    ('INCOME STATEMENT', 'D40'): calc_INCOME_STATEMENT_D40,
    ('INCOME STATEMENT', 'AE43'): calc_INCOME_STATEMENT_AE43,
    ('INCOME STATEMENT', 'AF43'): calc_INCOME_STATEMENT_AF43,
    ('INCOME STATEMENT', 'AG43'): calc_INCOME_STATEMENT_AG43,
    ('INCOME STATEMENT', 'AD45'): calc_INCOME_STATEMENT_AD45,
    ('INCOME STATEMENT', 'AE45'): calc_INCOME_STATEMENT_AE45,
    ('INCOME STATEMENT', 'AF45'): calc_INCOME_STATEMENT_AF45,
    ('INCOME STATEMENT', 'AG45'): calc_INCOME_STATEMENT_AG45,
    ('INCOME STATEMENT', 'C47'): calc_INCOME_STATEMENT_C47,
    ('INCOME STATEMENT', 'T47'): calc_INCOME_STATEMENT_T47,
    ('INCOME STATEMENT', 'B48'): calc_INCOME_STATEMENT_B48,
    ('INCOME STATEMENT', 'S48'): calc_INCOME_STATEMENT_S48,
    ('INCOME STATEMENT', 'W66'): calc_INCOME_STATEMENT_W66,
    ('CASH FOW STATEMENT', 'I8'): calc_CASH_FOW_STATEMENT_I8,
    ('CASH FOW STATEMENT', 'C9'): calc_CASH_FOW_STATEMENT_C9,
    ('CASH FOW STATEMENT', 'D9'): calc_CASH_FOW_STATEMENT_D9,
    ('CASH FOW STATEMENT', 'I9'): calc_CASH_FOW_STATEMENT_I9,
    ('CASH FOW STATEMENT', 'E11'): calc_CASH_FOW_STATEMENT_E11,
    ('CASH FOW STATEMENT', 'F11'): calc_CASH_FOW_STATEMENT_F11,
    ('CASH FOW STATEMENT', 'G11'): calc_CASH_FOW_STATEMENT_G11,
    ('CASH FOW STATEMENT', 'H11'): calc_CASH_FOW_STATEMENT_H11,
    ('CASH FOW STATEMENT', 'I11'): calc_CASH_FOW_STATEMENT_I11,
    ('CASH FOW STATEMENT', 'B13'): calc_CASH_FOW_STATEMENT_B13,
    ('CASH FOW STATEMENT', 'C13'): calc_CASH_FOW_STATEMENT_C13,
    ('CASH FOW STATEMENT', 'D13'): calc_CASH_FOW_STATEMENT_D13,
    ('CASH FOW STATEMENT', 'E13'): calc_CASH_FOW_STATEMENT_E13,
    ('CASH FOW STATEMENT', 'F13'): calc_CASH_FOW_STATEMENT_F13,
    ('CASH FOW STATEMENT', 'G13'): calc_CASH_FOW_STATEMENT_G13,
    ('CASH FOW STATEMENT', 'H13'): calc_CASH_FOW_STATEMENT_H13,
    ('CASH FOW STATEMENT', 'I13'): calc_CASH_FOW_STATEMENT_I13,
    ('CASH FOW STATEMENT', 'B17'): calc_CASH_FOW_STATEMENT_B17,
    ('CASH FOW STATEMENT', 'C17'): calc_CASH_FOW_STATEMENT_C17,
    ('CASH FOW STATEMENT', 'D17'): calc_CASH_FOW_STATEMENT_D17,
    ('CASH FOW STATEMENT', 'E17'): calc_CASH_FOW_STATEMENT_E17,
    ('CASH FOW STATEMENT', 'F17'): calc_CASH_FOW_STATEMENT_F17,
    ('CASH FOW STATEMENT', 'G17'): calc_CASH_FOW_STATEMENT_G17,
    ('CASH FOW STATEMENT', 'H17'): calc_CASH_FOW_STATEMENT_H17,
    ('CASH FOW STATEMENT', 'I17'): calc_CASH_FOW_STATEMENT_I17,
    ('CASH FOW STATEMENT', 'B19'): calc_CASH_FOW_STATEMENT_B19,
    ('CASH FOW STATEMENT', 'C19'): calc_CASH_FOW_STATEMENT_C19,
    ('CASH FOW STATEMENT', 'D19'): calc_CASH_FOW_STATEMENT_D19,
    ('CASH FOW STATEMENT', 'E19'): calc_CASH_FOW_STATEMENT_E19,
    ('CASH FOW STATEMENT', 'F19'): calc_CASH_FOW_STATEMENT_F19,
    ('CASH FOW STATEMENT', 'G19'): calc_CASH_FOW_STATEMENT_G19,
    ('CASH FOW STATEMENT', 'H19'): calc_CASH_FOW_STATEMENT_H19,
    ('CASH FOW STATEMENT', 'I19'): calc_CASH_FOW_STATEMENT_I19,
    ('CASH FOW STATEMENT', 'F20'): calc_CASH_FOW_STATEMENT_F20,
    ('CASH FOW STATEMENT', 'G20'): calc_CASH_FOW_STATEMENT_G20,
    ('CASH FOW STATEMENT', 'H20'): calc_CASH_FOW_STATEMENT_H20,
    ('CASH FOW STATEMENT', 'I20'): calc_CASH_FOW_STATEMENT_I20,
    ('CASH FOW STATEMENT', 'E22'): calc_CASH_FOW_STATEMENT_E22,
    ('BALANCESHEET', 'B10'): calc_BALANCESHEET_B10,
    ('BALANCESHEET', 'C10'): calc_BALANCESHEET_C10,
    ('BALANCESHEET', 'D10'): calc_BALANCESHEET_D10,
    ('BALANCESHEET', 'E10'): calc_BALANCESHEET_E10,
    ('BALANCESHEET', 'F10'): calc_BALANCESHEET_F10,
    ('BALANCESHEET', 'G10'): calc_BALANCESHEET_G10,
    ('BALANCESHEET', 'H10'): calc_BALANCESHEET_H10,
    ('BALANCESHEET', 'I10'): calc_BALANCESHEET_I10,
    ('BALANCESHEET', 'J10'): calc_BALANCESHEET_J10,
    ('BALANCESHEET', 'B19'): calc_BALANCESHEET_B19,
    ('BALANCESHEET', 'C19'): calc_BALANCESHEET_C19,
    ('BALANCESHEET', 'D19'): calc_BALANCESHEET_D19,
    ('BALANCESHEET', 'E19'): calc_BALANCESHEET_E19,
    ('BALANCESHEET', 'F19'): calc_BALANCESHEET_F19,
    ('BALANCESHEET', 'G19'): calc_BALANCESHEET_G19,
    ('BALANCESHEET', 'H19'): calc_BALANCESHEET_H19,
    ('BALANCESHEET', 'I19'): calc_BALANCESHEET_I19,
    ('BALANCESHEET', 'J19'): calc_BALANCESHEET_J19,
    ('BALANCESHEET', 'B30'): calc_BALANCESHEET_B30,
    ('BALANCESHEET', 'C30'): calc_BALANCESHEET_C30,
    ('BALANCESHEET', 'D30'): calc_BALANCESHEET_D30,
    ('BALANCESHEET', 'E30'): calc_BALANCESHEET_E30,
    ('BALANCESHEET', 'F30'): calc_BALANCESHEET_F30,
    ('BALANCESHEET', 'G30'): calc_BALANCESHEET_G30,
    ('BALANCESHEET', 'H30'): calc_BALANCESHEET_H30,
    ('BALANCESHEET', 'I30'): calc_BALANCESHEET_I30,
    ('BALANCESHEET', 'J30'): calc_BALANCESHEET_J30,
    ('BALANCESHEET', 'B36'): calc_BALANCESHEET_B36,
    ('BALANCESHEET', 'C36'): calc_BALANCESHEET_C36,
    ('BALANCESHEET', 'D36'): calc_BALANCESHEET_D36,
    ('BALANCESHEET', 'E36'): calc_BALANCESHEET_E36,
    ('BALANCESHEET', 'F36'): calc_BALANCESHEET_F36,
    ('BALANCESHEET', 'G36'): calc_BALANCESHEET_G36,
    ('BALANCESHEET', 'H36'): calc_BALANCESHEET_H36,
    ('BALANCESHEET', 'I36'): calc_BALANCESHEET_I36,
    ('BALANCESHEET', 'J36'): calc_BALANCESHEET_J36,
    ('BALANCESHEET', 'B44'): calc_BALANCESHEET_B44,
    ('BALANCESHEET', 'C44'): calc_BALANCESHEET_C44,
    ('BALANCESHEET', 'D44'): calc_BALANCESHEET_D44,
    ('BALANCESHEET', 'E44'): calc_BALANCESHEET_E44,
    ('BALANCESHEET', 'F44'): calc_BALANCESHEET_F44,
    ('BALANCESHEET', 'G44'): calc_BALANCESHEET_G44,
    ('BALANCESHEET', 'H44'): calc_BALANCESHEET_H44,
    ('BALANCESHEET', 'I44'): calc_BALANCESHEET_I44,
    ('BALANCESHEET', 'J44'): calc_BALANCESHEET_J44,
    ('BALANCESHEET', 'B52'): calc_BALANCESHEET_B52,
    ('BALANCESHEET', 'C52'): calc_BALANCESHEET_C52,
    ('BALANCESHEET', 'D52'): calc_BALANCESHEET_D52,
    ('BALANCESHEET', 'E52'): calc_BALANCESHEET_E52,
    ('Debt Schedule', 'B11'): calc_Debt_Schedule_B11,
    ('Debt Schedule', 'C11'): calc_Debt_Schedule_C11,
    ('Debt Schedule', 'D11'): calc_Debt_Schedule_D11,
    ('Debt Schedule', 'E11'): calc_Debt_Schedule_E11,
    ('Debt Schedule', 'F11'): calc_Debt_Schedule_F11,
    ('Debt Schedule', 'G11'): calc_Debt_Schedule_G11,
    ('Debt Schedule', 'H11'): calc_Debt_Schedule_H11,
    ('Debt Schedule', 'I11'): calc_Debt_Schedule_I11,
    ('Debt Schedule', 'B20'): calc_Debt_Schedule_B20,
    ('Debt Schedule', 'C20'): calc_Debt_Schedule_C20,
    ('Debt Schedule', 'D20'): calc_Debt_Schedule_D20,
    ('Debt Schedule', 'E20'): calc_Debt_Schedule_E20,
    ('Debt Schedule', 'F20'): calc_Debt_Schedule_F20,
    ('Debt Schedule', 'G20'): calc_Debt_Schedule_G20,
    ('Debt Schedule', 'H20'): calc_Debt_Schedule_H20,
    ('Debt Schedule', 'I20'): calc_Debt_Schedule_I20,
    ('Valuation', 'G9'): calc_Valuation_G9,
    ('Valuation', 'B10'): calc_Valuation_B10,
    ('Valuation', 'G11'): calc_Valuation_G11,
    ('Valuation', 'B19'): calc_Valuation_B19,
    ('Valuation', 'D19'): calc_Valuation_D19,
    ('Valuation', 'E19'): calc_Valuation_E19,
    ('Valuation', 'F19'): calc_Valuation_F19,
    ('Valuation', 'G19'): calc_Valuation_G19,
    ('Valuation', 'B22'): calc_Valuation_B22,
    ('Valuation', 'C24'): calc_Valuation_C24,
    ('Valuation', 'D24'): calc_Valuation_D24,
    ('Valuation', 'E24'): calc_Valuation_E24,
    ('Valuation', 'F24'): calc_Valuation_F24,
    ('Valuation', 'G24'): calc_Valuation_G24,
    ('Valuation', 'K37'): calc_Valuation_K37,
    ('Valuation', 'K53'): calc_Valuation_K53,
    ('Valuation', 'B54'): calc_Valuation_B54,
    ('Valuation', 'H64'): calc_Valuation_H64,
    ('Valuation', 'I64'): calc_Valuation_I64,
    ('Valuation', 'J64'): calc_Valuation_J64,
    ('Valuation', 'H74'): calc_Valuation_H74,
    ('Valuation', 'I74'): calc_Valuation_I74,
    ('Valuation', 'J74'): calc_Valuation_J74,
    ('Ratio Analysis', 'G14'): calc_Ratio_Analysis_G14,
    ('Ratio Analysis', 'H14'): calc_Ratio_Analysis_H14,
    ('Ratio Analysis', 'I14'): calc_Ratio_Analysis_I14,
    ('Ratio Analysis', 'J14'): calc_Ratio_Analysis_J14,
    ('Ratio Analysis', 'C20'): calc_Ratio_Analysis_C20,
    ('Ratio Analysis', 'D20'): calc_Ratio_Analysis_D20,
    ('Ratio Analysis', 'E20'): calc_Ratio_Analysis_E20,
    ('Ratio Analysis', 'F20'): calc_Ratio_Analysis_F20,
    ('Ratio Analysis', 'G20'): calc_Ratio_Analysis_G20,
    ('Ratio Analysis', 'H20'): calc_Ratio_Analysis_H20,
    ('Ratio Analysis', 'I20'): calc_Ratio_Analysis_I20,
    ('Ratio Analysis', 'J20'): calc_Ratio_Analysis_J20,
    ('Ratio Analysis', 'G21'): calc_Ratio_Analysis_G21,
    ('Ratio Analysis', 'H21'): calc_Ratio_Analysis_H21,
    ('Ratio Analysis', 'J21'): calc_Ratio_Analysis_J21,
    ('Ratio Analysis', 'E23'): calc_Ratio_Analysis_E23,
    ('Ratio Analysis', 'G23'): calc_Ratio_Analysis_G23,
    ('Ratio Analysis', 'H23'): calc_Ratio_Analysis_H23,
    ('Ratio Analysis', 'I23'): calc_Ratio_Analysis_I23,
    ('Ratio Analysis', 'J23'): calc_Ratio_Analysis_J23,
    ('Ratio Analysis', 'G24'): calc_Ratio_Analysis_G24,
    ('Ratio Analysis', 'H24'): calc_Ratio_Analysis_H24,
    ('Ratio Analysis', 'I24'): calc_Ratio_Analysis_I24,
    ('Ratio Analysis', 'J24'): calc_Ratio_Analysis_J24,
    ('PRESENTATION', 'F15'): calc_PRESENTATION_F15,
    ('PRESENTATION', 'K15'): calc_PRESENTATION_K15,
    ('PRESENTATION', 'P15'): calc_PRESENTATION_P15,
    ('PRESENTATION', 'T15'): calc_PRESENTATION_T15,
    ('PRESENTATION', 'B17'): calc_PRESENTATION_B17,
    ('PRESENTATION', 'C17'): calc_PRESENTATION_C17,
    ('PRESENTATION', 'D17'): calc_PRESENTATION_D17,
    ('PRESENTATION', 'E17'): calc_PRESENTATION_E17,
    ('PRESENTATION', 'G17'): calc_PRESENTATION_G17,
    ('PRESENTATION', 'H17'): calc_PRESENTATION_H17,
    ('PRESENTATION', 'I17'): calc_PRESENTATION_I17,
    ('PRESENTATION', 'J17'): calc_PRESENTATION_J17,
    ('PRESENTATION', 'L17'): calc_PRESENTATION_L17,
    ('PRESENTATION', 'M17'): calc_PRESENTATION_M17,
    ('PRESENTATION', 'N17'): calc_PRESENTATION_N17,
    ('PRESENTATION', 'O17'): calc_PRESENTATION_O17,
    ('PRESENTATION', 'Q17'): calc_PRESENTATION_Q17,
    ('PRESENTATION', 'R17'): calc_PRESENTATION_R17,
    ('PRESENTATION', 'S17'): calc_PRESENTATION_S17,
    ('PRESENTATION', 'V17'): calc_PRESENTATION_V17,
    ('PRESENTATION', 'W17'): calc_PRESENTATION_W17,
    ('PRESENTATION', 'X17'): calc_PRESENTATION_X17,
    ('PRESENTATION', 'Y17'): calc_PRESENTATION_Y17,
    ('PRESENTATION', 'F25'): calc_PRESENTATION_F25,
    ('PRESENTATION', 'M25'): calc_PRESENTATION_M25,
    ('PRESENTATION', 'T25'): calc_PRESENTATION_T25,
    ('PRESENTATION', 'AA25'): calc_PRESENTATION_AA25,
    ('PRESENTATION', 'B32'): calc_PRESENTATION_B32,
    ('PRESENTATION', 'C32'): calc_PRESENTATION_C32,
    ('PRESENTATION', 'E32'): calc_PRESENTATION_E32,
    ('PRESENTATION', 'G32'): calc_PRESENTATION_G32,
    ('PRESENTATION', 'I32'): calc_PRESENTATION_I32,
    ('PRESENTATION', 'J32'): calc_PRESENTATION_J32,
    ('PRESENTATION', 'L32'): calc_PRESENTATION_L32,
    ('PRESENTATION', 'N32'): calc_PRESENTATION_N32,
    ('PRESENTATION', 'P32'): calc_PRESENTATION_P32,
    ('PRESENTATION', 'Q32'): calc_PRESENTATION_Q32,
    ('PRESENTATION', 'S32'): calc_PRESENTATION_S32,
    ('PRESENTATION', 'U32'): calc_PRESENTATION_U32,
    ('PRESENTATION', 'W32'): calc_PRESENTATION_W32,
    ('PRESENTATION', 'X32'): calc_PRESENTATION_X32,
    ('PRESENTATION', 'Z32'): calc_PRESENTATION_Z32,
    ('PRESENTATION', 'AB32'): calc_PRESENTATION_AB32,
    ('PRESENTATION', 'AD32'): calc_PRESENTATION_AD32,
    ('PRESENTATION', 'AE32'): calc_PRESENTATION_AE32,
    ('PRESENTATION', 'AF32'): calc_PRESENTATION_AF32,
    ('PRESENTATION', 'AG32'): calc_PRESENTATION_AG32,
    ('PRESENTATION', 'F27'): calc_PRESENTATION_F27,
    ('PRESENTATION', 'M27'): calc_PRESENTATION_M27,
    ('PRESENTATION', 'T27'): calc_PRESENTATION_T27,
    ('PRESENTATION', 'AA27'): calc_PRESENTATION_AA27,
    ('PRESENTATION', 'F28'): calc_PRESENTATION_F28,
    ('PRESENTATION', 'M28'): calc_PRESENTATION_M28,
    ('PRESENTATION', 'T28'): calc_PRESENTATION_T28,
    ('PRESENTATION', 'AA28'): calc_PRESENTATION_AA28,
    ('PRESENTATION', 'F29'): calc_PRESENTATION_F29,
    ('PRESENTATION', 'M29'): calc_PRESENTATION_M29,
    ('PRESENTATION', 'T29'): calc_PRESENTATION_T29,
    ('PRESENTATION', 'AA29'): calc_PRESENTATION_AA29,
    ('PRESENTATION', 'F30'): calc_PRESENTATION_F30,
    ('PRESENTATION', 'M30'): calc_PRESENTATION_M30,
    ('PRESENTATION', 'T30'): calc_PRESENTATION_T30,
    ('PRESENTATION', 'AA30'): calc_PRESENTATION_AA30,
    ('PRESENTATION', 'D26'): calc_PRESENTATION_D26,
    ('PRESENTATION', 'F31'): calc_PRESENTATION_F31,
    ('PRESENTATION', 'M31'): calc_PRESENTATION_M31,
    ('PRESENTATION', 'R26'): calc_PRESENTATION_R26,
    ('PRESENTATION', 'T31'): calc_PRESENTATION_T31,
    ('PRESENTATION', 'Y26'): calc_PRESENTATION_Y26,
    ('PRESENTATION', 'AA31'): calc_PRESENTATION_AA31,
    ('PRESENTATION', 'K33'): calc_PRESENTATION_K33,
    ('PRESENTATION', 'F34'): calc_PRESENTATION_F34,
    ('PRESENTATION', 'M34'): calc_PRESENTATION_M34,
    ('PRESENTATION', 'T34'): calc_PRESENTATION_T34,
    ('PRESENTATION', 'AA34'): calc_PRESENTATION_AA34,
    ('PRESENTATION', 'F35'): calc_PRESENTATION_F35,
    ('PRESENTATION', 'M35'): calc_PRESENTATION_M35,
    ('PRESENTATION', 'K36'): calc_PRESENTATION_K36,
    ('PRESENTATION', 'T35'): calc_PRESENTATION_T35,
    ('PRESENTATION', 'AA35'): calc_PRESENTATION_AA35,
    ('PRESENTATION', 'F37'): calc_PRESENTATION_F37,
    ('PRESENTATION', 'M37'): calc_PRESENTATION_M37,
    ('PRESENTATION', 'T37'): calc_PRESENTATION_T37,
    ('PRESENTATION', 'AA37'): calc_PRESENTATION_AA37,
    ('PRESENTATION', 'T39'): calc_PRESENTATION_T39,
    ('PRESENTATION', 'AA39'): calc_PRESENTATION_AA39,
    ('PRESENTATION', 'F41'): calc_PRESENTATION_F41,
    ('PRESENTATION', 'M41'): calc_PRESENTATION_M41,
    ('PRESENTATION', 'T41'): calc_PRESENTATION_T41,
    ('PRESENTATION', 'AA41'): calc_PRESENTATION_AA41,
    ('PRESENTATION', 'F42'): calc_PRESENTATION_F42,
    ('PRESENTATION', 'M42'): calc_PRESENTATION_M42,
    ('PRESENTATION', 'T42'): calc_PRESENTATION_T42,
    ('PRESENTATION', 'AA42'): calc_PRESENTATION_AA42,
    ('PRESENTATION', 'F43'): calc_PRESENTATION_F43,
    ('PRESENTATION', 'M43'): calc_PRESENTATION_M43,
    ('PRESENTATION', 'T43'): calc_PRESENTATION_T43,
    ('PRESENTATION', 'AA43'): calc_PRESENTATION_AA43,
    ('Segment Revenue Model', 'K24'): calc_Segment_Revenue_Model_K24,
    ('Segment Revenue Model', 'U8'): calc_Segment_Revenue_Model_U8,
    ('Segment Revenue Model', 'P24'): calc_Segment_Revenue_Model_P24,
    ('Segment Revenue Model', 'K25'): calc_Segment_Revenue_Model_K25,
    ('Segment Revenue Model', 'U9'): calc_Segment_Revenue_Model_U9,
    ('Segment Revenue Model', 'P25'): calc_Segment_Revenue_Model_P25,
    ('Segment Revenue Model', 'K26'): calc_Segment_Revenue_Model_K26,
    ('Segment Revenue Model', 'U10'): calc_Segment_Revenue_Model_U10,
    ('Segment Revenue Model', 'P26'): calc_Segment_Revenue_Model_P26,
    ('Segment Revenue Model', 'K27'): calc_Segment_Revenue_Model_K27,
    ('Segment Revenue Model', 'U11'): calc_Segment_Revenue_Model_U11,
    ('Segment Revenue Model', 'P27'): calc_Segment_Revenue_Model_P27,
    ('Segment Revenue Model', 'K28'): calc_Segment_Revenue_Model_K28,
    ('Segment Revenue Model', 'P28'): calc_Segment_Revenue_Model_P28,
    ('Segment Revenue Model', 'T12'): calc_Segment_Revenue_Model_T12,
    ('Segment Revenue Model', 'V12'): calc_Segment_Revenue_Model_V12,
    ('Segment Revenue Model', 'K29'): calc_Segment_Revenue_Model_K29,
    ('Segment Revenue Model', 'U13'): calc_Segment_Revenue_Model_U13,
    ('Segment Revenue Model', 'P29'): calc_Segment_Revenue_Model_P29,
    ('Segment Revenue Model', 'F15'): calc_Segment_Revenue_Model_F15,
    ('Segment Revenue Model', 'K15'): calc_Segment_Revenue_Model_K15,
    ('Segment Revenue Model', 'K30'): calc_Segment_Revenue_Model_K30,
    ('Segment Revenue Model', 'U14'): calc_Segment_Revenue_Model_U14,
    ('Segment Revenue Model', 'P15'): calc_Segment_Revenue_Model_P15,
    ('Segment Revenue Model', 'P30'): calc_Segment_Revenue_Model_P30,
    ('Segment Revenue Model', 'B17'): calc_Segment_Revenue_Model_B17,
    ('Segment Revenue Model', 'C17'): calc_Segment_Revenue_Model_C17,
    ('Segment Revenue Model', 'D17'): calc_Segment_Revenue_Model_D17,
    ('Segment Revenue Model', 'E17'): calc_Segment_Revenue_Model_E17,
    ('Segment Revenue Model', 'G17'): calc_Segment_Revenue_Model_G17,
    ('Segment Revenue Model', 'G31'): calc_Segment_Revenue_Model_G31,
    ('Segment Revenue Model', 'H17'): calc_Segment_Revenue_Model_H17,
    ('Segment Revenue Model', 'H31'): calc_Segment_Revenue_Model_H31,
    ('Segment Revenue Model', 'I17'): calc_Segment_Revenue_Model_I17,
    ('Segment Revenue Model', 'I31'): calc_Segment_Revenue_Model_I31,
    ('Segment Revenue Model', 'J17'): calc_Segment_Revenue_Model_J17,
    ('Segment Revenue Model', 'J31'): calc_Segment_Revenue_Model_J31,
    ('Segment Revenue Model', 'L17'): calc_Segment_Revenue_Model_L17,
    ('Segment Revenue Model', 'L31'): calc_Segment_Revenue_Model_L31,
    ('Segment Revenue Model', 'M17'): calc_Segment_Revenue_Model_M17,
    ('Segment Revenue Model', 'M31'): calc_Segment_Revenue_Model_M31,
    ('Segment Revenue Model', 'N17'): calc_Segment_Revenue_Model_N17,
    ('Segment Revenue Model', 'N31'): calc_Segment_Revenue_Model_N31,
    ('Segment Revenue Model', 'O17'): calc_Segment_Revenue_Model_O17,
    ('Segment Revenue Model', 'O31'): calc_Segment_Revenue_Model_O31,
    ('Segment Revenue Model', 'Q17'): calc_Segment_Revenue_Model_Q17,
    ('Segment Revenue Model', 'Q31'): calc_Segment_Revenue_Model_Q31,
    ('Segment Revenue Model', 'R17'): calc_Segment_Revenue_Model_R17,
    ('Segment Revenue Model', 'R31'): calc_Segment_Revenue_Model_R31,
    ('Segment Revenue Model', 'S17'): calc_Segment_Revenue_Model_S17,
    ('Segment Revenue Model', 'S31'): calc_Segment_Revenue_Model_S31,
    ('Segment Revenue Model', 'K32'): calc_Segment_Revenue_Model_K32,
    ('Segment Revenue Model', 'P32'): calc_Segment_Revenue_Model_P32,
    ('Segment Revenue Model', 'V16'): calc_Segment_Revenue_Model_V16,
    ('Segment Revenue Model', 'F58'): calc_Segment_Revenue_Model_F58,
    ('Segment Revenue Model', 'K58'): calc_Segment_Revenue_Model_K58,
    ('Segment Revenue Model', 'P58'): calc_Segment_Revenue_Model_P58,
    ('Segment Revenue Model', 'U58'): calc_Segment_Revenue_Model_U58,
    ('Segment Revenue Model', 'B61'): calc_Segment_Revenue_Model_B61,
    ('Segment Revenue Model', 'C61'): calc_Segment_Revenue_Model_C61,
    ('Segment Revenue Model', 'D61'): calc_Segment_Revenue_Model_D61,
    ('Segment Revenue Model', 'E61'): calc_Segment_Revenue_Model_E61,
    ('Segment Revenue Model', 'G61'): calc_Segment_Revenue_Model_G61,
    ('Segment Revenue Model', 'H61'): calc_Segment_Revenue_Model_H61,
    ('Segment Revenue Model', 'I61'): calc_Segment_Revenue_Model_I61,
    ('Segment Revenue Model', 'J61'): calc_Segment_Revenue_Model_J61,
    ('Segment Revenue Model', 'M61'): calc_Segment_Revenue_Model_M61,
    ('Segment Revenue Model', 'N61'): calc_Segment_Revenue_Model_N61,
    ('Segment Revenue Model', 'O61'): calc_Segment_Revenue_Model_O61,
    ('Segment Revenue Model', 'Q61'): calc_Segment_Revenue_Model_Q61,
    ('Segment Revenue Model', 'R61'): calc_Segment_Revenue_Model_R61,
    ('Segment Revenue Model', 'S61'): calc_Segment_Revenue_Model_S61,
    ('Segment Revenue Model', 'T61'): calc_Segment_Revenue_Model_T61,
    ('Segment Revenue Model', 'V61'): calc_Segment_Revenue_Model_V61,
    ('Segment Revenue Model', 'W61'): calc_Segment_Revenue_Model_W61,
    ('Segment Revenue Model', 'X61'): calc_Segment_Revenue_Model_X61,
    ('Segment Revenue Model', 'Y61'): calc_Segment_Revenue_Model_Y61,
    ('Segment Revenue Model', 'F72'): calc_Segment_Revenue_Model_F72,
    ('Segment Revenue Model', 'K72'): calc_Segment_Revenue_Model_K72,
    ('Segment Revenue Model', 'P72'): calc_Segment_Revenue_Model_P72,
    ('Segment Revenue Model', 'B74'): calc_Segment_Revenue_Model_B74,
    ('Segment Revenue Model', 'C74'): calc_Segment_Revenue_Model_C74,
    ('Segment Revenue Model', 'D74'): calc_Segment_Revenue_Model_D74,
    ('Segment Revenue Model', 'E74'): calc_Segment_Revenue_Model_E74,
    ('Segment Revenue Model', 'G74'): calc_Segment_Revenue_Model_G74,
    ('Segment Revenue Model', 'H74'): calc_Segment_Revenue_Model_H74,
    ('Segment Revenue Model', 'I74'): calc_Segment_Revenue_Model_I74,
    ('Segment Revenue Model', 'J74'): calc_Segment_Revenue_Model_J74,
    ('Segment Revenue Model', 'L74'): calc_Segment_Revenue_Model_L74,
    ('Segment Revenue Model', 'M74'): calc_Segment_Revenue_Model_M74,
    ('Segment Revenue Model', 'N74'): calc_Segment_Revenue_Model_N74,
    ('Segment Revenue Model', 'O74'): calc_Segment_Revenue_Model_O74,
    ('Segment Revenue Model', 'Q74'): calc_Segment_Revenue_Model_Q74,
    ('Segment Revenue Model', 'R74'): calc_Segment_Revenue_Model_R74,
    ('Segment Revenue Model', 'S74'): calc_Segment_Revenue_Model_S74,
    ('Segment Revenue Model', 'U72'): calc_Segment_Revenue_Model_U72,
    ('Segment Revenue Model', 'T74'): calc_Segment_Revenue_Model_T74,
    ('INCOME STATEMENT', 'F6'): calc_INCOME_STATEMENT_F6,
    ('INCOME STATEMENT', 'D7'): calc_INCOME_STATEMENT_D7,
    ('INCOME STATEMENT', 'K8'): calc_INCOME_STATEMENT_K8,
    ('INCOME STATEMENT', 'M6'): calc_INCOME_STATEMENT_M6,
    ('INCOME STATEMENT', 'T6'): calc_INCOME_STATEMENT_T6,
    ('INCOME STATEMENT', 'R7'): calc_INCOME_STATEMENT_R7,
    ('INCOME STATEMENT', 'R8'): calc_INCOME_STATEMENT_R8,
    ('INCOME STATEMENT', 'AA6'): calc_INCOME_STATEMENT_AA6,
    ('INCOME STATEMENT', 'Y8'): calc_INCOME_STATEMENT_Y8,
    ('Ratio Analysis', 'I21'): calc_Ratio_Analysis_I21,
    ('INCOME STATEMENT', 'B21'): calc_INCOME_STATEMENT_B21,
    ('INCOME STATEMENT', 'C21'): calc_INCOME_STATEMENT_C21,
    ('INCOME STATEMENT', 'E21'): calc_INCOME_STATEMENT_E21,
    ('INCOME STATEMENT', 'G21'): calc_INCOME_STATEMENT_G21,
    ('INCOME STATEMENT', 'I21'): calc_INCOME_STATEMENT_I21,
    ('INCOME STATEMENT', 'J21'): calc_INCOME_STATEMENT_J21,
    ('INCOME STATEMENT', 'L21'): calc_INCOME_STATEMENT_L21,
    ('INCOME STATEMENT', 'N21'): calc_INCOME_STATEMENT_N21,
    ('INCOME STATEMENT', 'P21'): calc_INCOME_STATEMENT_P21,
    ('INCOME STATEMENT', 'Q21'): calc_INCOME_STATEMENT_Q21,
    ('INCOME STATEMENT', 'S21'): calc_INCOME_STATEMENT_S21,
    ('INCOME STATEMENT', 'U21'): calc_INCOME_STATEMENT_U21,
    ('INCOME STATEMENT', 'W21'): calc_INCOME_STATEMENT_W21,
    ('INCOME STATEMENT', 'X21'): calc_INCOME_STATEMENT_X21,
    ('INCOME STATEMENT', 'Z21'): calc_INCOME_STATEMENT_Z21,
    ('INCOME STATEMENT', 'AB21'): calc_INCOME_STATEMENT_AB21,
    ('INCOME STATEMENT', 'F11'): calc_INCOME_STATEMENT_F11,
    ('INCOME STATEMENT', 'D12'): calc_INCOME_STATEMENT_D12,
    ('INCOME STATEMENT', 'M11'): calc_INCOME_STATEMENT_M11,
    ('INCOME STATEMENT', 'K12'): calc_INCOME_STATEMENT_K12,
    ('INCOME STATEMENT', 'T11'): calc_INCOME_STATEMENT_T11,
    ('INCOME STATEMENT', 'R12'): calc_INCOME_STATEMENT_R12,
    ('INCOME STATEMENT', 'AA11'): calc_INCOME_STATEMENT_AA11,
    ('INCOME STATEMENT', 'Y12'): calc_INCOME_STATEMENT_Y12,
    ('INCOME STATEMENT', 'F13'): calc_INCOME_STATEMENT_F13,
    ('INCOME STATEMENT', 'D14'): calc_INCOME_STATEMENT_D14,
    ('INCOME STATEMENT', 'M13'): calc_INCOME_STATEMENT_M13,
    ('INCOME STATEMENT', 'K14'): calc_INCOME_STATEMENT_K14,
    ('INCOME STATEMENT', 'T13'): calc_INCOME_STATEMENT_T13,
    ('INCOME STATEMENT', 'R14'): calc_INCOME_STATEMENT_R14,
    ('INCOME STATEMENT', 'AA13'): calc_INCOME_STATEMENT_AA13,
    ('INCOME STATEMENT', 'Y14'): calc_INCOME_STATEMENT_Y14,
    ('INCOME STATEMENT', 'F15'): calc_INCOME_STATEMENT_F15,
    ('INCOME STATEMENT', 'D16'): calc_INCOME_STATEMENT_D16,
    ('INCOME STATEMENT', 'M15'): calc_INCOME_STATEMENT_M15,
    ('INCOME STATEMENT', 'K16'): calc_INCOME_STATEMENT_K16,
    ('INCOME STATEMENT', 'T15'): calc_INCOME_STATEMENT_T15,
    ('INCOME STATEMENT', 'R16'): calc_INCOME_STATEMENT_R16,
    ('INCOME STATEMENT', 'AA15'): calc_INCOME_STATEMENT_AA15,
    ('INCOME STATEMENT', 'Y16'): calc_INCOME_STATEMENT_Y16,
    ('INCOME STATEMENT', 'F17'): calc_INCOME_STATEMENT_F17,
    ('INCOME STATEMENT', 'D18'): calc_INCOME_STATEMENT_D18,
    ('INCOME STATEMENT', 'M17'): calc_INCOME_STATEMENT_M17,
    ('INCOME STATEMENT', 'K18'): calc_INCOME_STATEMENT_K18,
    ('INCOME STATEMENT', 'T17'): calc_INCOME_STATEMENT_T17,
    ('INCOME STATEMENT', 'R18'): calc_INCOME_STATEMENT_R18,
    ('INCOME STATEMENT', 'AA17'): calc_INCOME_STATEMENT_AA17,
    ('INCOME STATEMENT', 'Y18'): calc_INCOME_STATEMENT_Y18,
    ('INCOME STATEMENT', 'D10'): calc_INCOME_STATEMENT_D10,
    ('INCOME STATEMENT', 'F19'): calc_INCOME_STATEMENT_F19,
    ('INCOME STATEMENT', 'D20'): calc_INCOME_STATEMENT_D20,
    ('INCOME STATEMENT', 'M19'): calc_INCOME_STATEMENT_M19,
    ('INCOME STATEMENT', 'K20'): calc_INCOME_STATEMENT_K20,
    ('INCOME STATEMENT', 'R10'): calc_INCOME_STATEMENT_R10,
    ('INCOME STATEMENT', 'T19'): calc_INCOME_STATEMENT_T19,
    ('INCOME STATEMENT', 'R20'): calc_INCOME_STATEMENT_R20,
    ('INCOME STATEMENT', 'Y10'): calc_INCOME_STATEMENT_Y10,
    ('INCOME STATEMENT', 'AA19'): calc_INCOME_STATEMENT_AA19,
    ('INCOME STATEMENT', 'Y20'): calc_INCOME_STATEMENT_Y20,
    ('INCOME STATEMENT', 'AD10'): calc_INCOME_STATEMENT_AD10,
    ('INCOME STATEMENT', 'AE10'): calc_INCOME_STATEMENT_AE10,
    ('INCOME STATEMENT', 'AF10'): calc_INCOME_STATEMENT_AF10,
    ('INCOME STATEMENT', 'AG10'): calc_INCOME_STATEMENT_AG10,
    ('INCOME STATEMENT', 'K25'): calc_INCOME_STATEMENT_K25,
    ('INCOME STATEMENT', 'F26'): calc_INCOME_STATEMENT_F26,
    ('INCOME STATEMENT', 'M26'): calc_INCOME_STATEMENT_M26,
    ('INCOME STATEMENT', 'T26'): calc_INCOME_STATEMENT_T26,
    ('INCOME STATEMENT', 'AA26'): calc_INCOME_STATEMENT_AA26,
    ('INCOME STATEMENT', 'F27'): calc_INCOME_STATEMENT_F27,
    ('INCOME STATEMENT', 'M27'): calc_INCOME_STATEMENT_M27,
    ('INCOME STATEMENT', 'K28'): calc_INCOME_STATEMENT_K28,
    ('INCOME STATEMENT', 'T27'): calc_INCOME_STATEMENT_T27,
    ('INCOME STATEMENT', 'AA27'): calc_INCOME_STATEMENT_AA27,
    ('INCOME STATEMENT', 'F29'): calc_INCOME_STATEMENT_F29,
    ('INCOME STATEMENT', 'M29'): calc_INCOME_STATEMENT_M29,
    ('INCOME STATEMENT', 'T29'): calc_INCOME_STATEMENT_T29,
    ('INCOME STATEMENT', 'AA29'): calc_INCOME_STATEMENT_AA29,
    ('INCOME STATEMENT', 'T31'): calc_INCOME_STATEMENT_T31,
    ('INCOME STATEMENT', 'AA31'): calc_INCOME_STATEMENT_AA31,
    ('INCOME STATEMENT', 'F33'): calc_INCOME_STATEMENT_F33,
    ('INCOME STATEMENT', 'M33'): calc_INCOME_STATEMENT_M33,
    ('INCOME STATEMENT', 'T33'): calc_INCOME_STATEMENT_T33,
    ('INCOME STATEMENT', 'AA33'): calc_INCOME_STATEMENT_AA33,
    ('INCOME STATEMENT', 'F34'): calc_INCOME_STATEMENT_F34,
    ('INCOME STATEMENT', 'M34'): calc_INCOME_STATEMENT_M34,
    ('INCOME STATEMENT', 'T34'): calc_INCOME_STATEMENT_T34,
    ('INCOME STATEMENT', 'AA34'): calc_INCOME_STATEMENT_AA34,
    ('INCOME STATEMENT', 'F35'): calc_INCOME_STATEMENT_F35,
    ('INCOME STATEMENT', 'M35'): calc_INCOME_STATEMENT_M35,
    ('INCOME STATEMENT', 'T35'): calc_INCOME_STATEMENT_T35,
    ('INCOME STATEMENT', 'AA35'): calc_INCOME_STATEMENT_AA35,
    ('INCOME STATEMENT', 'D47'): calc_INCOME_STATEMENT_D47,
    ('INCOME STATEMENT', 'C48'): calc_INCOME_STATEMENT_C48,
    ('INCOME STATEMENT', 'U47'): calc_INCOME_STATEMENT_U47,
    ('INCOME STATEMENT', 'T48'): calc_INCOME_STATEMENT_T48,
    ('CASH FOW STATEMENT', 'I10'): calc_CASH_FOW_STATEMENT_I10,
    ('Valuation', 'B27'): calc_Valuation_B27,
    ('Valuation', 'C27'): calc_Valuation_C27,
    ('Valuation', 'D27'): calc_Valuation_D27,
    ('Valuation', 'E27'): calc_Valuation_E27,
    ('Valuation', 'F27'): calc_Valuation_F27,
    ('Valuation', 'G27'): calc_Valuation_G27,
    ('Valuation', 'D25'): calc_Valuation_D25,
    ('Valuation', 'E25'): calc_Valuation_E25,
    ('Valuation', 'F25'): calc_Valuation_F25,
    ('Valuation', 'G25'): calc_Valuation_G25,
    ('BALANCESHEET', 'B8'): calc_BALANCESHEET_B8,
    ('BALANCESHEET', 'C8'): calc_BALANCESHEET_C8,
    ('BALANCESHEET', 'D8'): calc_BALANCESHEET_D8,
    ('BALANCESHEET', 'E8'): calc_BALANCESHEET_E8,
    ('BALANCESHEET', 'F8'): calc_BALANCESHEET_F8,
    ('BALANCESHEET', 'G8'): calc_BALANCESHEET_G8,
    ('BALANCESHEET', 'H8'): calc_BALANCESHEET_H8,
    ('BALANCESHEET', 'I8'): calc_BALANCESHEET_I8,
    ('BALANCESHEET', 'J8'): calc_BALANCESHEET_J8,
    ('BALANCESHEET', 'B24'): calc_BALANCESHEET_B24,
    ('CASH FOW STATEMENT', 'B21'): calc_CASH_FOW_STATEMENT_B21,
    ('BALANCESHEET', 'C24'): calc_BALANCESHEET_C24,
    ('CASH FOW STATEMENT', 'C21'): calc_CASH_FOW_STATEMENT_C21,
    ('BALANCESHEET', 'D24'): calc_BALANCESHEET_D24,
    ('CASH FOW STATEMENT', 'D21'): calc_CASH_FOW_STATEMENT_D21,
    ('BALANCESHEET', 'E24'): calc_BALANCESHEET_E24,
    ('Valuation', 'G10'): calc_Valuation_G10,
    ('CASH FOW STATEMENT', 'E21'): calc_CASH_FOW_STATEMENT_E21,
    ('BALANCESHEET', 'F24'): calc_BALANCESHEET_F24,
    ('CASH FOW STATEMENT', 'F21'): calc_CASH_FOW_STATEMENT_F21,
    ('BALANCESHEET', 'G24'): calc_BALANCESHEET_G24,
    ('Ratio Analysis', 'G15'): calc_Ratio_Analysis_G15,
    ('CASH FOW STATEMENT', 'G21'): calc_CASH_FOW_STATEMENT_G21,
    ('BALANCESHEET', 'H24'): calc_BALANCESHEET_H24,
    ('Ratio Analysis', 'H15'): calc_Ratio_Analysis_H15,
    ('CASH FOW STATEMENT', 'H21'): calc_CASH_FOW_STATEMENT_H21,
    ('BALANCESHEET', 'I24'): calc_BALANCESHEET_I24,
    ('Ratio Analysis', 'I15'): calc_Ratio_Analysis_I15,
    ('CASH FOW STATEMENT', 'I21'): calc_CASH_FOW_STATEMENT_I21,
    ('BALANCESHEET', 'J24'): calc_BALANCESHEET_J24,
    ('Ratio Analysis', 'J15'): calc_Ratio_Analysis_J15,
    ('BALANCESHEET', 'B27'): calc_BALANCESHEET_B27,
    ('BALANCESHEET', 'C27'): calc_BALANCESHEET_C27,
    ('BALANCESHEET', 'D27'): calc_BALANCESHEET_D27,
    ('BALANCESHEET', 'E27'): calc_BALANCESHEET_E27,
    ('BALANCESHEET', 'F27'): calc_BALANCESHEET_F27,
    ('BALANCESHEET', 'G27'): calc_BALANCESHEET_G27,
    ('BALANCESHEET', 'H27'): calc_BALANCESHEET_H27,
    ('BALANCESHEET', 'I27'): calc_BALANCESHEET_I27,
    ('BALANCESHEET', 'J27'): calc_BALANCESHEET_J27,
    ('Ratio Analysis', 'F9'): calc_Ratio_Analysis_F9,
    ('Ratio Analysis', 'G9'): calc_Ratio_Analysis_G9,
    ('Ratio Analysis', 'H9'): calc_Ratio_Analysis_H9,
    ('Ratio Analysis', 'I9'): calc_Ratio_Analysis_I9,
    ('Ratio Analysis', 'J9'): calc_Ratio_Analysis_J9,
    ('BALANCESHEET', 'B43'): calc_BALANCESHEET_B43,
    ('BALANCESHEET', 'B58'): calc_BALANCESHEET_B58,
    ('BALANCESHEET', 'C43'): calc_BALANCESHEET_C43,
    ('BALANCESHEET', 'C58'): calc_BALANCESHEET_C58,
    ('BALANCESHEET', 'D43'): calc_BALANCESHEET_D43,
    ('BALANCESHEET', 'D58'): calc_BALANCESHEET_D58,
    ('BALANCESHEET', 'E43'): calc_BALANCESHEET_E43,
    ('BALANCESHEET', 'E58'): calc_BALANCESHEET_E58,
    ('Debt Schedule', 'B21'): calc_Debt_Schedule_B21,
    ('Debt Schedule', 'C21'): calc_Debt_Schedule_C21,
    ('Debt Schedule', 'D21'): calc_Debt_Schedule_D21,
    ('Debt Schedule', 'E21'): calc_Debt_Schedule_E21,
    ('Debt Schedule', 'F21'): calc_Debt_Schedule_F21,
    ('Debt Schedule', 'G21'): calc_Debt_Schedule_G21,
    ('Debt Schedule', 'H21'): calc_Debt_Schedule_H21,
    ('Debt Schedule', 'I21'): calc_Debt_Schedule_I21,
    ('Valuation', 'G54'): calc_Valuation_G54,
    ('Valuation', 'E37'): calc_Valuation_E37,
    ('Valuation', 'D40'): calc_Valuation_D40,
    ('Valuation', 'D45'): calc_Valuation_D45,
    ('Valuation', 'E40'): calc_Valuation_E40,
    ('Valuation', 'E45'): calc_Valuation_E45,
    ('Valuation', 'F40'): calc_Valuation_F40,
    ('Valuation', 'G40'): calc_Valuation_G40,
    ('PRESENTATION', 'G65'): calc_PRESENTATION_G65,
    ('PRESENTATION', 'H65'): calc_PRESENTATION_H65,
    ('PRESENTATION', 'I65'): calc_PRESENTATION_I65,
    ('PRESENTATION', 'J65'): calc_PRESENTATION_J65,
    ('PRESENTATION', 'I58'): calc_PRESENTATION_I58,
    ('PRESENTATION', 'J58'): calc_PRESENTATION_J58,
    ('PRESENTATION', 'K17'): calc_PRESENTATION_K17,
    ('PRESENTATION', 'P17'): calc_PRESENTATION_P17,
    ('PRESENTATION', 'U15'): calc_PRESENTATION_U15,
    ('PRESENTATION', 'T17'): calc_PRESENTATION_T17,
    ('PRESENTATION', 'F17'): calc_PRESENTATION_F17,
    ('PRESENTATION', 'O25'): calc_PRESENTATION_O25,
    ('PRESENTATION', 'AC25'): calc_PRESENTATION_AC25,
    ('PRESENTATION', 'B33'): calc_PRESENTATION_B33,
    ('PRESENTATION', 'B36'): calc_PRESENTATION_B36,
    ('PRESENTATION', 'C33'): calc_PRESENTATION_C33,
    ('PRESENTATION', 'C36'): calc_PRESENTATION_C36,
    ('PRESENTATION', 'E33'): calc_PRESENTATION_E33,
    ('PRESENTATION', 'E36'): calc_PRESENTATION_E36,
    ('PRESENTATION', 'G33'): calc_PRESENTATION_G33,
    ('PRESENTATION', 'G36'): calc_PRESENTATION_G36,
    ('PRESENTATION', 'I33'): calc_PRESENTATION_I33,
    ('PRESENTATION', 'I36'): calc_PRESENTATION_I36,
    ('PRESENTATION', 'J33'): calc_PRESENTATION_J33,
    ('PRESENTATION', 'J36'): calc_PRESENTATION_J36,
    ('PRESENTATION', 'L33'): calc_PRESENTATION_L33,
    ('PRESENTATION', 'L36'): calc_PRESENTATION_L36,
    ('PRESENTATION', 'N33'): calc_PRESENTATION_N33,
    ('PRESENTATION', 'N36'): calc_PRESENTATION_N36,
    ('PRESENTATION', 'P33'): calc_PRESENTATION_P33,
    ('PRESENTATION', 'P36'): calc_PRESENTATION_P36,
    ('PRESENTATION', 'Q33'): calc_PRESENTATION_Q33,
    ('PRESENTATION', 'Q36'): calc_PRESENTATION_Q36,
    ('PRESENTATION', 'S33'): calc_PRESENTATION_S33,
    ('PRESENTATION', 'S36'): calc_PRESENTATION_S36,
    ('PRESENTATION', 'U33'): calc_PRESENTATION_U33,
    ('PRESENTATION', 'U36'): calc_PRESENTATION_U36,
    ('PRESENTATION', 'W33'): calc_PRESENTATION_W33,
    ('PRESENTATION', 'W36'): calc_PRESENTATION_W36,
    ('PRESENTATION', 'X33'): calc_PRESENTATION_X33,
    ('PRESENTATION', 'X36'): calc_PRESENTATION_X36,
    ('PRESENTATION', 'Z33'): calc_PRESENTATION_Z33,
    ('PRESENTATION', 'Z36'): calc_PRESENTATION_Z36,
    ('PRESENTATION', 'AB33'): calc_PRESENTATION_AB33,
    ('PRESENTATION', 'AB36'): calc_PRESENTATION_AB36,
    ('PRESENTATION', 'AD33'): calc_PRESENTATION_AD33,
    ('PRESENTATION', 'AD36'): calc_PRESENTATION_AD36,
    ('PRESENTATION', 'AE33'): calc_PRESENTATION_AE33,
    ('PRESENTATION', 'AE36'): calc_PRESENTATION_AE36,
    ('PRESENTATION', 'AF33'): calc_PRESENTATION_AF33,
    ('PRESENTATION', 'AF36'): calc_PRESENTATION_AF36,
    ('PRESENTATION', 'AG33'): calc_PRESENTATION_AG33,
    ('PRESENTATION', 'AG36'): calc_PRESENTATION_AG36,
    ('PRESENTATION', 'O27'): calc_PRESENTATION_O27,
    ('PRESENTATION', 'AC27'): calc_PRESENTATION_AC27,
    ('PRESENTATION', 'AC28'): calc_PRESENTATION_AC28,
    ('PRESENTATION', 'H29'): calc_PRESENTATION_H29,
    ('PRESENTATION', 'O29'): calc_PRESENTATION_O29,
    ('PRESENTATION', 'V29'): calc_PRESENTATION_V29,
    ('PRESENTATION', 'AC29'): calc_PRESENTATION_AC29,
    ('PRESENTATION', 'O30'): calc_PRESENTATION_O30,
    ('PRESENTATION', 'AC30'): calc_PRESENTATION_AC30,
    ('PRESENTATION', 'D32'): calc_PRESENTATION_D32,
    ('PRESENTATION', 'F26'): calc_PRESENTATION_F26,
    ('PRESENTATION', 'H31'): calc_PRESENTATION_H31,
    ('PRESENTATION', 'M26'): calc_PRESENTATION_M26,
    ('PRESENTATION', 'R32'): calc_PRESENTATION_R32,
    ('PRESENTATION', 'T26'): calc_PRESENTATION_T26,
    ('PRESENTATION', 'V31'): calc_PRESENTATION_V31,
    ('PRESENTATION', 'Y32'): calc_PRESENTATION_Y32,
    ('PRESENTATION', 'AA26'): calc_PRESENTATION_AA26,
    ('PRESENTATION', 'AC31'): calc_PRESENTATION_AC31,
    ('PRESENTATION', 'H34'): calc_PRESENTATION_H34,
    ('PRESENTATION', 'O34'): calc_PRESENTATION_O34,
    ('PRESENTATION', 'V34'): calc_PRESENTATION_V34,
    ('PRESENTATION', 'AC34'): calc_PRESENTATION_AC34,
    ('PRESENTATION', 'H35'): calc_PRESENTATION_H35,
    ('PRESENTATION', 'O35'): calc_PRESENTATION_O35,
    ('PRESENTATION', 'K38'): calc_PRESENTATION_K38,
    ('PRESENTATION', 'V35'): calc_PRESENTATION_V35,
    ('PRESENTATION', 'H37'): calc_PRESENTATION_H37,
    ('PRESENTATION', 'O37'): calc_PRESENTATION_O37,
    ('PRESENTATION', 'AC37'): calc_PRESENTATION_AC37,
    ('PRESENTATION', 'AC41'): calc_PRESENTATION_AC41,
    ('PRESENTATION', 'H42'): calc_PRESENTATION_H42,
    ('PRESENTATION', 'AC42'): calc_PRESENTATION_AC42,
    ('PRESENTATION', 'H43'): calc_PRESENTATION_H43,
    ('Segment Revenue Model', 'T8'): calc_Segment_Revenue_Model_T8,
    ('Segment Revenue Model', 'V8'): calc_Segment_Revenue_Model_V8,
    ('Segment Revenue Model', 'T9'): calc_Segment_Revenue_Model_T9,
    ('Segment Revenue Model', 'V9'): calc_Segment_Revenue_Model_V9,
    ('Segment Revenue Model', 'T10'): calc_Segment_Revenue_Model_T10,
    ('Segment Revenue Model', 'V10'): calc_Segment_Revenue_Model_V10,
    ('Segment Revenue Model', 'T11'): calc_Segment_Revenue_Model_T11,
    ('Segment Revenue Model', 'V11'): calc_Segment_Revenue_Model_V11,
    ('Segment Revenue Model', 'T28'): calc_Segment_Revenue_Model_T28,
    ('Segment Revenue Model', 'W12'): calc_Segment_Revenue_Model_W12,
    ('Segment Revenue Model', 'T13'): calc_Segment_Revenue_Model_T13,
    ('Segment Revenue Model', 'V13'): calc_Segment_Revenue_Model_V13,
    ('Segment Revenue Model', 'K17'): calc_Segment_Revenue_Model_K17,
    ('Segment Revenue Model', 'K31'): calc_Segment_Revenue_Model_K31,
    ('Segment Revenue Model', 'T14'): calc_Segment_Revenue_Model_T14,
    ('Segment Revenue Model', 'V14'): calc_Segment_Revenue_Model_V14,
    ('Segment Revenue Model', 'P17'): calc_Segment_Revenue_Model_P17,
    ('Segment Revenue Model', 'P31'): calc_Segment_Revenue_Model_P31,
    ('Segment Revenue Model', 'F17'): calc_Segment_Revenue_Model_F17,
    ('Segment Revenue Model', 'G33'): calc_Segment_Revenue_Model_G33,
    ('Segment Revenue Model', 'G37'): calc_Segment_Revenue_Model_G37,
    ('Segment Revenue Model', 'G38'): calc_Segment_Revenue_Model_G38,
    ('Segment Revenue Model', 'G39'): calc_Segment_Revenue_Model_G39,
    ('Segment Revenue Model', 'G40'): calc_Segment_Revenue_Model_G40,
    ('Segment Revenue Model', 'G41'): calc_Segment_Revenue_Model_G41,
    ('Segment Revenue Model', 'G42'): calc_Segment_Revenue_Model_G42,
    ('Segment Revenue Model', 'G43'): calc_Segment_Revenue_Model_G43,
    ('Segment Revenue Model', 'G44'): calc_Segment_Revenue_Model_G44,
    ('Segment Revenue Model', 'G45'): calc_Segment_Revenue_Model_G45,
    ('Segment Revenue Model', 'G46'): calc_Segment_Revenue_Model_G46,
    ('Segment Revenue Model', 'H33'): calc_Segment_Revenue_Model_H33,
    ('Segment Revenue Model', 'H37'): calc_Segment_Revenue_Model_H37,
    ('Segment Revenue Model', 'H38'): calc_Segment_Revenue_Model_H38,
    ('Segment Revenue Model', 'H39'): calc_Segment_Revenue_Model_H39,
    ('Segment Revenue Model', 'H40'): calc_Segment_Revenue_Model_H40,
    ('Segment Revenue Model', 'H41'): calc_Segment_Revenue_Model_H41,
    ('Segment Revenue Model', 'H42'): calc_Segment_Revenue_Model_H42,
    ('Segment Revenue Model', 'H43'): calc_Segment_Revenue_Model_H43,
    ('Segment Revenue Model', 'H44'): calc_Segment_Revenue_Model_H44,
    ('Segment Revenue Model', 'H45'): calc_Segment_Revenue_Model_H45,
    ('Segment Revenue Model', 'H46'): calc_Segment_Revenue_Model_H46,
    ('Segment Revenue Model', 'I33'): calc_Segment_Revenue_Model_I33,
    ('Segment Revenue Model', 'I37'): calc_Segment_Revenue_Model_I37,
    ('Segment Revenue Model', 'I38'): calc_Segment_Revenue_Model_I38,
    ('Segment Revenue Model', 'I39'): calc_Segment_Revenue_Model_I39,
    ('Segment Revenue Model', 'I40'): calc_Segment_Revenue_Model_I40,
    ('Segment Revenue Model', 'I41'): calc_Segment_Revenue_Model_I41,
    ('Segment Revenue Model', 'I42'): calc_Segment_Revenue_Model_I42,
    ('Segment Revenue Model', 'I43'): calc_Segment_Revenue_Model_I43,
    ('Segment Revenue Model', 'I44'): calc_Segment_Revenue_Model_I44,
    ('Segment Revenue Model', 'I45'): calc_Segment_Revenue_Model_I45,
    ('Segment Revenue Model', 'I46'): calc_Segment_Revenue_Model_I46,
    ('Segment Revenue Model', 'J33'): calc_Segment_Revenue_Model_J33,
    ('Segment Revenue Model', 'J37'): calc_Segment_Revenue_Model_J37,
    ('Segment Revenue Model', 'J38'): calc_Segment_Revenue_Model_J38,
    ('Segment Revenue Model', 'J39'): calc_Segment_Revenue_Model_J39,
    ('Segment Revenue Model', 'J40'): calc_Segment_Revenue_Model_J40,
    ('Segment Revenue Model', 'J41'): calc_Segment_Revenue_Model_J41,
    ('Segment Revenue Model', 'J42'): calc_Segment_Revenue_Model_J42,
    ('Segment Revenue Model', 'J43'): calc_Segment_Revenue_Model_J43,
    ('Segment Revenue Model', 'J44'): calc_Segment_Revenue_Model_J44,
    ('Segment Revenue Model', 'J45'): calc_Segment_Revenue_Model_J45,
    ('Segment Revenue Model', 'J46'): calc_Segment_Revenue_Model_J46,
    ('Segment Revenue Model', 'L33'): calc_Segment_Revenue_Model_L33,
    ('Segment Revenue Model', 'L37'): calc_Segment_Revenue_Model_L37,
    ('Segment Revenue Model', 'L38'): calc_Segment_Revenue_Model_L38,
    ('Segment Revenue Model', 'L39'): calc_Segment_Revenue_Model_L39,
    ('Segment Revenue Model', 'L40'): calc_Segment_Revenue_Model_L40,
    ('Segment Revenue Model', 'L41'): calc_Segment_Revenue_Model_L41,
    ('Segment Revenue Model', 'L42'): calc_Segment_Revenue_Model_L42,
    ('Segment Revenue Model', 'L43'): calc_Segment_Revenue_Model_L43,
    ('Segment Revenue Model', 'L44'): calc_Segment_Revenue_Model_L44,
    ('Segment Revenue Model', 'L45'): calc_Segment_Revenue_Model_L45,
    ('Segment Revenue Model', 'L46'): calc_Segment_Revenue_Model_L46,
    ('Segment Revenue Model', 'M33'): calc_Segment_Revenue_Model_M33,
    ('Segment Revenue Model', 'M37'): calc_Segment_Revenue_Model_M37,
    ('Segment Revenue Model', 'M38'): calc_Segment_Revenue_Model_M38,
    ('Segment Revenue Model', 'M39'): calc_Segment_Revenue_Model_M39,
    ('Segment Revenue Model', 'M40'): calc_Segment_Revenue_Model_M40,
    ('Segment Revenue Model', 'M41'): calc_Segment_Revenue_Model_M41,
    ('Segment Revenue Model', 'M42'): calc_Segment_Revenue_Model_M42,
    ('Segment Revenue Model', 'M43'): calc_Segment_Revenue_Model_M43,
    ('Segment Revenue Model', 'M44'): calc_Segment_Revenue_Model_M44,
    ('Segment Revenue Model', 'M45'): calc_Segment_Revenue_Model_M45,
    ('Segment Revenue Model', 'M46'): calc_Segment_Revenue_Model_M46,
    ('Segment Revenue Model', 'N33'): calc_Segment_Revenue_Model_N33,
    ('Segment Revenue Model', 'N37'): calc_Segment_Revenue_Model_N37,
    ('Segment Revenue Model', 'N38'): calc_Segment_Revenue_Model_N38,
    ('Segment Revenue Model', 'N39'): calc_Segment_Revenue_Model_N39,
    ('Segment Revenue Model', 'N40'): calc_Segment_Revenue_Model_N40,
    ('Segment Revenue Model', 'N41'): calc_Segment_Revenue_Model_N41,
    ('Segment Revenue Model', 'N42'): calc_Segment_Revenue_Model_N42,
    ('Segment Revenue Model', 'N43'): calc_Segment_Revenue_Model_N43,
    ('Segment Revenue Model', 'N44'): calc_Segment_Revenue_Model_N44,
    ('Segment Revenue Model', 'N45'): calc_Segment_Revenue_Model_N45,
    ('Segment Revenue Model', 'N46'): calc_Segment_Revenue_Model_N46,
    ('Segment Revenue Model', 'O33'): calc_Segment_Revenue_Model_O33,
    ('Segment Revenue Model', 'O37'): calc_Segment_Revenue_Model_O37,
    ('Segment Revenue Model', 'O38'): calc_Segment_Revenue_Model_O38,
    ('Segment Revenue Model', 'O39'): calc_Segment_Revenue_Model_O39,
    ('Segment Revenue Model', 'O40'): calc_Segment_Revenue_Model_O40,
    ('Segment Revenue Model', 'O41'): calc_Segment_Revenue_Model_O41,
    ('Segment Revenue Model', 'O42'): calc_Segment_Revenue_Model_O42,
    ('Segment Revenue Model', 'O43'): calc_Segment_Revenue_Model_O43,
    ('Segment Revenue Model', 'O44'): calc_Segment_Revenue_Model_O44,
    ('Segment Revenue Model', 'O45'): calc_Segment_Revenue_Model_O45,
    ('Segment Revenue Model', 'O46'): calc_Segment_Revenue_Model_O46,
    ('Segment Revenue Model', 'Q33'): calc_Segment_Revenue_Model_Q33,
    ('Segment Revenue Model', 'Q37'): calc_Segment_Revenue_Model_Q37,
    ('Segment Revenue Model', 'Q38'): calc_Segment_Revenue_Model_Q38,
    ('Segment Revenue Model', 'Q39'): calc_Segment_Revenue_Model_Q39,
    ('Segment Revenue Model', 'Q40'): calc_Segment_Revenue_Model_Q40,
    ('Segment Revenue Model', 'Q41'): calc_Segment_Revenue_Model_Q41,
    ('Segment Revenue Model', 'Q42'): calc_Segment_Revenue_Model_Q42,
    ('Segment Revenue Model', 'Q43'): calc_Segment_Revenue_Model_Q43,
    ('Segment Revenue Model', 'Q44'): calc_Segment_Revenue_Model_Q44,
    ('Segment Revenue Model', 'Q45'): calc_Segment_Revenue_Model_Q45,
    ('Segment Revenue Model', 'Q46'): calc_Segment_Revenue_Model_Q46,
    ('Segment Revenue Model', 'R33'): calc_Segment_Revenue_Model_R33,
    ('Segment Revenue Model', 'R37'): calc_Segment_Revenue_Model_R37,
    ('Segment Revenue Model', 'R38'): calc_Segment_Revenue_Model_R38,
    ('Segment Revenue Model', 'R39'): calc_Segment_Revenue_Model_R39,
    ('Segment Revenue Model', 'R40'): calc_Segment_Revenue_Model_R40,
    ('Segment Revenue Model', 'R41'): calc_Segment_Revenue_Model_R41,
    ('Segment Revenue Model', 'R42'): calc_Segment_Revenue_Model_R42,
    ('Segment Revenue Model', 'R43'): calc_Segment_Revenue_Model_R43,
    ('Segment Revenue Model', 'R44'): calc_Segment_Revenue_Model_R44,
    ('Segment Revenue Model', 'R45'): calc_Segment_Revenue_Model_R45,
    ('Segment Revenue Model', 'R46'): calc_Segment_Revenue_Model_R46,
    ('Segment Revenue Model', 'S33'): calc_Segment_Revenue_Model_S33,
    ('Segment Revenue Model', 'S37'): calc_Segment_Revenue_Model_S37,
    ('Segment Revenue Model', 'S38'): calc_Segment_Revenue_Model_S38,
    ('Segment Revenue Model', 'S39'): calc_Segment_Revenue_Model_S39,
    ('Segment Revenue Model', 'S40'): calc_Segment_Revenue_Model_S40,
    ('Segment Revenue Model', 'S41'): calc_Segment_Revenue_Model_S41,
    ('Segment Revenue Model', 'S42'): calc_Segment_Revenue_Model_S42,
    ('Segment Revenue Model', 'S43'): calc_Segment_Revenue_Model_S43,
    ('Segment Revenue Model', 'S44'): calc_Segment_Revenue_Model_S44,
    ('Segment Revenue Model', 'S45'): calc_Segment_Revenue_Model_S45,
    ('Segment Revenue Model', 'S46'): calc_Segment_Revenue_Model_S46,
    ('Segment Revenue Model', 'W16'): calc_Segment_Revenue_Model_W16,
    ('Segment Revenue Model', 'F61'): calc_Segment_Revenue_Model_F61,
    ('Segment Revenue Model', 'P61'): calc_Segment_Revenue_Model_P61,
    ('Segment Revenue Model', 'U61'): calc_Segment_Revenue_Model_U61,
    ('Segment Revenue Model', 'K61'): calc_Segment_Revenue_Model_K61,
    ('Segment Revenue Model', 'K74'): calc_Segment_Revenue_Model_K74,
    ('Segment Revenue Model', 'P74'): calc_Segment_Revenue_Model_P74,
    ('Segment Revenue Model', 'F74'): calc_Segment_Revenue_Model_F74,
    ('Segment Revenue Model', 'U74'): calc_Segment_Revenue_Model_U74,
    ('INCOME STATEMENT', 'F7'): calc_INCOME_STATEMENT_F7,
    ('INCOME STATEMENT', 'O6'): calc_INCOME_STATEMENT_O6,
    ('INCOME STATEMENT', 'M8'): calc_INCOME_STATEMENT_M8,
    ('INCOME STATEMENT', 'T7'): calc_INCOME_STATEMENT_T7,
    ('INCOME STATEMENT', 'T8'): calc_INCOME_STATEMENT_T8,
    ('INCOME STATEMENT', 'AC6'): calc_INCOME_STATEMENT_AC6,
    ('INCOME STATEMENT', 'AA8'): calc_INCOME_STATEMENT_AA8,
    ('INCOME STATEMENT', 'B25'): calc_INCOME_STATEMENT_B25,
    ('INCOME STATEMENT', 'B28'): calc_INCOME_STATEMENT_B28,
    ('INCOME STATEMENT', 'C24'): calc_INCOME_STATEMENT_C24,
    ('INCOME STATEMENT', 'C25'): calc_INCOME_STATEMENT_C25,
    ('INCOME STATEMENT', 'C28'): calc_INCOME_STATEMENT_C28,
    ('INCOME STATEMENT', 'E24'): calc_INCOME_STATEMENT_E24,
    ('INCOME STATEMENT', 'E25'): calc_INCOME_STATEMENT_E25,
    ('INCOME STATEMENT', 'E28'): calc_INCOME_STATEMENT_E28,
    ('INCOME STATEMENT', 'G24'): calc_INCOME_STATEMENT_G24,
    ('INCOME STATEMENT', 'G25'): calc_INCOME_STATEMENT_G25,
    ('INCOME STATEMENT', 'G28'): calc_INCOME_STATEMENT_G28,
    ('INCOME STATEMENT', 'I23'): calc_INCOME_STATEMENT_I23,
    ('INCOME STATEMENT', 'I24'): calc_INCOME_STATEMENT_I24,
    ('INCOME STATEMENT', 'I25'): calc_INCOME_STATEMENT_I25,
    ('INCOME STATEMENT', 'I28'): calc_INCOME_STATEMENT_I28,
    ('INCOME STATEMENT', 'J23'): calc_INCOME_STATEMENT_J23,
    ('INCOME STATEMENT', 'J24'): calc_INCOME_STATEMENT_J24,
    ('INCOME STATEMENT', 'J25'): calc_INCOME_STATEMENT_J25,
    ('INCOME STATEMENT', 'J28'): calc_INCOME_STATEMENT_J28,
    ('INCOME STATEMENT', 'L23'): calc_INCOME_STATEMENT_L23,
    ('INCOME STATEMENT', 'L24'): calc_INCOME_STATEMENT_L24,
    ('INCOME STATEMENT', 'L25'): calc_INCOME_STATEMENT_L25,
    ('INCOME STATEMENT', 'L28'): calc_INCOME_STATEMENT_L28,
    ('INCOME STATEMENT', 'N23'): calc_INCOME_STATEMENT_N23,
    ('INCOME STATEMENT', 'N24'): calc_INCOME_STATEMENT_N24,
    ('INCOME STATEMENT', 'N25'): calc_INCOME_STATEMENT_N25,
    ('INCOME STATEMENT', 'N28'): calc_INCOME_STATEMENT_N28,
    ('INCOME STATEMENT', 'P23'): calc_INCOME_STATEMENT_P23,
    ('INCOME STATEMENT', 'P24'): calc_INCOME_STATEMENT_P24,
    ('INCOME STATEMENT', 'P25'): calc_INCOME_STATEMENT_P25,
    ('INCOME STATEMENT', 'P28'): calc_INCOME_STATEMENT_P28,
    ('INCOME STATEMENT', 'Q23'): calc_INCOME_STATEMENT_Q23,
    ('INCOME STATEMENT', 'Q24'): calc_INCOME_STATEMENT_Q24,
    ('INCOME STATEMENT', 'Q25'): calc_INCOME_STATEMENT_Q25,
    ('INCOME STATEMENT', 'Q28'): calc_INCOME_STATEMENT_Q28,
    ('INCOME STATEMENT', 'S23'): calc_INCOME_STATEMENT_S23,
    ('INCOME STATEMENT', 'S24'): calc_INCOME_STATEMENT_S24,
    ('INCOME STATEMENT', 'S25'): calc_INCOME_STATEMENT_S25,
    ('INCOME STATEMENT', 'S28'): calc_INCOME_STATEMENT_S28,
    ('INCOME STATEMENT', 'U23'): calc_INCOME_STATEMENT_U23,
    ('INCOME STATEMENT', 'U24'): calc_INCOME_STATEMENT_U24,
    ('INCOME STATEMENT', 'U25'): calc_INCOME_STATEMENT_U25,
    ('INCOME STATEMENT', 'U28'): calc_INCOME_STATEMENT_U28,
    ('INCOME STATEMENT', 'W23'): calc_INCOME_STATEMENT_W23,
    ('INCOME STATEMENT', 'W24'): calc_INCOME_STATEMENT_W24,
    ('INCOME STATEMENT', 'W25'): calc_INCOME_STATEMENT_W25,
    ('INCOME STATEMENT', 'W28'): calc_INCOME_STATEMENT_W28,
    ('INCOME STATEMENT', 'X23'): calc_INCOME_STATEMENT_X23,
    ('INCOME STATEMENT', 'X24'): calc_INCOME_STATEMENT_X24,
    ('INCOME STATEMENT', 'X25'): calc_INCOME_STATEMENT_X25,
    ('INCOME STATEMENT', 'X28'): calc_INCOME_STATEMENT_X28,
    ('INCOME STATEMENT', 'Z23'): calc_INCOME_STATEMENT_Z23,
    ('INCOME STATEMENT', 'Z24'): calc_INCOME_STATEMENT_Z24,
    ('INCOME STATEMENT', 'Z25'): calc_INCOME_STATEMENT_Z25,
    ('INCOME STATEMENT', 'Z28'): calc_INCOME_STATEMENT_Z28,
    ('INCOME STATEMENT', 'AB23'): calc_INCOME_STATEMENT_AB23,
    ('INCOME STATEMENT', 'AB24'): calc_INCOME_STATEMENT_AB24,
    ('INCOME STATEMENT', 'AB25'): calc_INCOME_STATEMENT_AB25,
    ('INCOME STATEMENT', 'AB28'): calc_INCOME_STATEMENT_AB28,
    ('INCOME STATEMENT', 'F12'): calc_INCOME_STATEMENT_F12,
    ('INCOME STATEMENT', 'O11'): calc_INCOME_STATEMENT_O11,
    ('INCOME STATEMENT', 'M12'): calc_INCOME_STATEMENT_M12,
    ('INCOME STATEMENT', 'T12'): calc_INCOME_STATEMENT_T12,
    ('INCOME STATEMENT', 'AC11'): calc_INCOME_STATEMENT_AC11,
    ('INCOME STATEMENT', 'AA12'): calc_INCOME_STATEMENT_AA12,
    ('INCOME STATEMENT', 'F14'): calc_INCOME_STATEMENT_F14,
    ('INCOME STATEMENT', 'M14'): calc_INCOME_STATEMENT_M14,
    ('INCOME STATEMENT', 'T14'): calc_INCOME_STATEMENT_T14,
    ('INCOME STATEMENT', 'AC13'): calc_INCOME_STATEMENT_AC13,
    ('INCOME STATEMENT', 'AA14'): calc_INCOME_STATEMENT_AA14,
    ('INCOME STATEMENT', 'H15'): calc_INCOME_STATEMENT_H15,
    ('INCOME STATEMENT', 'F16'): calc_INCOME_STATEMENT_F16,
    ('INCOME STATEMENT', 'O15'): calc_INCOME_STATEMENT_O15,
    ('INCOME STATEMENT', 'M16'): calc_INCOME_STATEMENT_M16,
    ('INCOME STATEMENT', 'V15'): calc_INCOME_STATEMENT_V15,
    ('INCOME STATEMENT', 'T16'): calc_INCOME_STATEMENT_T16,
    ('INCOME STATEMENT', 'AC15'): calc_INCOME_STATEMENT_AC15,
    ('INCOME STATEMENT', 'AA16'): calc_INCOME_STATEMENT_AA16,
    ('INCOME STATEMENT', 'F18'): calc_INCOME_STATEMENT_F18,
    ('INCOME STATEMENT', 'O17'): calc_INCOME_STATEMENT_O17,
    ('INCOME STATEMENT', 'M18'): calc_INCOME_STATEMENT_M18,
    ('INCOME STATEMENT', 'T18'): calc_INCOME_STATEMENT_T18,
    ('INCOME STATEMENT', 'AC17'): calc_INCOME_STATEMENT_AC17,
    ('INCOME STATEMENT', 'AA18'): calc_INCOME_STATEMENT_AA18,
    ('INCOME STATEMENT', 'D21'): calc_INCOME_STATEMENT_D21,
    ('INCOME STATEMENT', 'F10'): calc_INCOME_STATEMENT_F10,
    ('INCOME STATEMENT', 'H19'): calc_INCOME_STATEMENT_H19,
    ('INCOME STATEMENT', 'F20'): calc_INCOME_STATEMENT_F20,
    ('INCOME STATEMENT', 'M10'): calc_INCOME_STATEMENT_M10,
    ('INCOME STATEMENT', 'M20'): calc_INCOME_STATEMENT_M20,
    ('INCOME STATEMENT', 'R21'): calc_INCOME_STATEMENT_R21,
    ('INCOME STATEMENT', 'T10'): calc_INCOME_STATEMENT_T10,
    ('INCOME STATEMENT', 'V19'): calc_INCOME_STATEMENT_V19,
    ('INCOME STATEMENT', 'T20'): calc_INCOME_STATEMENT_T20,
    ('INCOME STATEMENT', 'Y21'): calc_INCOME_STATEMENT_Y21,
    ('INCOME STATEMENT', 'AA10'): calc_INCOME_STATEMENT_AA10,
    ('INCOME STATEMENT', 'AC19'): calc_INCOME_STATEMENT_AC19,
    ('INCOME STATEMENT', 'AA20'): calc_INCOME_STATEMENT_AA20,
    ('PRESENTATION', 'G64'): calc_PRESENTATION_G64,
    ('INCOME STATEMENT', 'AD21'): calc_INCOME_STATEMENT_AD21,
    ('PRESENTATION', 'H64'): calc_PRESENTATION_H64,
    ('INCOME STATEMENT', 'AE21'): calc_INCOME_STATEMENT_AE21,
    ('PRESENTATION', 'I64'): calc_PRESENTATION_I64,
    ('INCOME STATEMENT', 'AF21'): calc_INCOME_STATEMENT_AF21,
    ('PRESENTATION', 'J64'): calc_PRESENTATION_J64,
    ('INCOME STATEMENT', 'AG21'): calc_INCOME_STATEMENT_AG21,
    ('INCOME STATEMENT', 'H26'): calc_INCOME_STATEMENT_H26,
    ('INCOME STATEMENT', 'O26'): calc_INCOME_STATEMENT_O26,
    ('INCOME STATEMENT', 'V26'): calc_INCOME_STATEMENT_V26,
    ('INCOME STATEMENT', 'AC26'): calc_INCOME_STATEMENT_AC26,
    ('INCOME STATEMENT', 'H27'): calc_INCOME_STATEMENT_H27,
    ('INCOME STATEMENT', 'O27'): calc_INCOME_STATEMENT_O27,
    ('INCOME STATEMENT', 'K30'): calc_INCOME_STATEMENT_K30,
    ('INCOME STATEMENT', 'V27'): calc_INCOME_STATEMENT_V27,
    ('INCOME STATEMENT', 'H29'): calc_INCOME_STATEMENT_H29,
    ('INCOME STATEMENT', 'O29'): calc_INCOME_STATEMENT_O29,
    ('INCOME STATEMENT', 'AC29'): calc_INCOME_STATEMENT_AC29,
    ('INCOME STATEMENT', 'AC33'): calc_INCOME_STATEMENT_AC33,
    ('INCOME STATEMENT', 'H34'): calc_INCOME_STATEMENT_H34,
    ('INCOME STATEMENT', 'AC34'): calc_INCOME_STATEMENT_AC34,
    ('INCOME STATEMENT', 'H35'): calc_INCOME_STATEMENT_H35,
    ('INCOME STATEMENT', 'E47'): calc_INCOME_STATEMENT_E47,
    ('INCOME STATEMENT', 'D48'): calc_INCOME_STATEMENT_D48,
    ('INCOME STATEMENT', 'V47'): calc_INCOME_STATEMENT_V47,
    ('INCOME STATEMENT', 'U48'): calc_INCOME_STATEMENT_U48,
    ('CASH FOW STATEMENT', 'I12'): calc_CASH_FOW_STATEMENT_I12,
    ('Valuation', 'B46'): calc_Valuation_B46,
    ('Valuation', 'D46'): calc_Valuation_D46,
    ('Valuation', 'E46'): calc_Valuation_E46,
    ('Valuation', 'H25'): calc_Valuation_H25,
    ('Ratio Analysis', 'B8'): calc_Ratio_Analysis_B8,
    ('Ratio Analysis', 'C8'): calc_Ratio_Analysis_C8,
    ('Ratio Analysis', 'D8'): calc_Ratio_Analysis_D8,
    ('Ratio Analysis', 'E8'): calc_Ratio_Analysis_E8,
    ('Ratio Analysis', 'F8'): calc_Ratio_Analysis_F8,
    ('Ratio Analysis', 'G8'): calc_Ratio_Analysis_G8,
    ('Ratio Analysis', 'H8'): calc_Ratio_Analysis_H8,
    ('Ratio Analysis', 'I8'): calc_Ratio_Analysis_I8,
    ('Ratio Analysis', 'J8'): calc_Ratio_Analysis_J8,
    ('Ratio Analysis', 'B7'): calc_Ratio_Analysis_B7,
    ('Ratio Analysis', 'C7'): calc_Ratio_Analysis_C7,
    ('Ratio Analysis', 'D7'): calc_Ratio_Analysis_D7,
    ('Ratio Analysis', 'E7'): calc_Ratio_Analysis_E7,
    ('Valuation', 'G12'): calc_Valuation_G12,
    ('Valuation', 'G53'): calc_Valuation_G53,
    ('Ratio Analysis', 'F7'): calc_Ratio_Analysis_F7,
    ('Ratio Analysis', 'G7'): calc_Ratio_Analysis_G7,
    ('Ratio Analysis', 'H7'): calc_Ratio_Analysis_H7,
    ('Ratio Analysis', 'I7'): calc_Ratio_Analysis_I7,
    ('Ratio Analysis', 'J7'): calc_Ratio_Analysis_J7,
    ('Ratio Analysis', 'B17'): calc_Ratio_Analysis_B17,
    ('Ratio Analysis', 'C10'): calc_Ratio_Analysis_C10,
    ('Ratio Analysis', 'C17'): calc_Ratio_Analysis_C17,
    ('Ratio Analysis', 'D17'): calc_Ratio_Analysis_D17,
    ('Ratio Analysis', 'E10'): calc_Ratio_Analysis_E10,
    ('Ratio Analysis', 'E17'): calc_Ratio_Analysis_E17,
    ('BALANCESHEET', 'F60'): calc_BALANCESHEET_F60,
    ('Ratio Analysis', 'F17'): calc_Ratio_Analysis_F17,
    ('BALANCESHEET', 'G60'): calc_BALANCESHEET_G60,
    ('Ratio Analysis', 'G10'): calc_Ratio_Analysis_G10,
    ('Ratio Analysis', 'G17'): calc_Ratio_Analysis_G17,
    ('BALANCESHEET', 'H60'): calc_BALANCESHEET_H60,
    ('Ratio Analysis', 'H10'): calc_Ratio_Analysis_H10,
    ('Ratio Analysis', 'H17'): calc_Ratio_Analysis_H17,
    ('BALANCESHEET', 'I60'): calc_BALANCESHEET_I60,
    ('Ratio Analysis', 'I10'): calc_Ratio_Analysis_I10,
    ('Ratio Analysis', 'I17'): calc_Ratio_Analysis_I17,
    ('BALANCESHEET', 'J60'): calc_BALANCESHEET_J60,
    ('Ratio Analysis', 'J10'): calc_Ratio_Analysis_J10,
    ('Ratio Analysis', 'J17'): calc_Ratio_Analysis_J17,
    ('Ratio Analysis', 'B9'): calc_Ratio_Analysis_B9,
    ('BALANCESHEET', 'B60'): calc_BALANCESHEET_B60,
    ('CASH FOW STATEMENT', 'B20'): calc_CASH_FOW_STATEMENT_B20,
    ('Ratio Analysis', 'C9'): calc_Ratio_Analysis_C9,
    ('BALANCESHEET', 'C60'): calc_BALANCESHEET_C60,
    ('Ratio Analysis', 'C24'): calc_Ratio_Analysis_C24,
    ('CASH FOW STATEMENT', 'C20'): calc_CASH_FOW_STATEMENT_C20,
    ('Ratio Analysis', 'D9'): calc_Ratio_Analysis_D9,
    ('BALANCESHEET', 'D60'): calc_BALANCESHEET_D60,
    ('CASH FOW STATEMENT', 'D20'): calc_CASH_FOW_STATEMENT_D20,
    ('CASH FOW STATEMENT', 'E20'): calc_CASH_FOW_STATEMENT_E20,
    ('Ratio Analysis', 'E9'): calc_Ratio_Analysis_E9,
    ('BALANCESHEET', 'E60'): calc_BALANCESHEET_E60,
    ('Ratio Analysis', 'E24'): calc_Ratio_Analysis_E24,
    ('Valuation', 'F37'): calc_Valuation_F37,
    ('Valuation', 'F45'): calc_Valuation_F45,
    ('PRESENTATION', 'U17'): calc_PRESENTATION_U17,
    ('PRESENTATION', 'B38'): calc_PRESENTATION_B38,
    ('PRESENTATION', 'C38'): calc_PRESENTATION_C38,
    ('PRESENTATION', 'E38'): calc_PRESENTATION_E38,
    ('PRESENTATION', 'G38'): calc_PRESENTATION_G38,
    ('PRESENTATION', 'I38'): calc_PRESENTATION_I38,
    ('PRESENTATION', 'J38'): calc_PRESENTATION_J38,
    ('PRESENTATION', 'L38'): calc_PRESENTATION_L38,
    ('PRESENTATION', 'N38'): calc_PRESENTATION_N38,
    ('PRESENTATION', 'P38'): calc_PRESENTATION_P38,
    ('PRESENTATION', 'Q38'): calc_PRESENTATION_Q38,
    ('PRESENTATION', 'S38'): calc_PRESENTATION_S38,
    ('PRESENTATION', 'U38'): calc_PRESENTATION_U38,
    ('PRESENTATION', 'W38'): calc_PRESENTATION_W38,
    ('PRESENTATION', 'X38'): calc_PRESENTATION_X38,
    ('PRESENTATION', 'Z38'): calc_PRESENTATION_Z38,
    ('PRESENTATION', 'AB38'): calc_PRESENTATION_AB38,
    ('PRESENTATION', 'AD38'): calc_PRESENTATION_AD38,
    ('PRESENTATION', 'AE38'): calc_PRESENTATION_AE38,
    ('PRESENTATION', 'AF38'): calc_PRESENTATION_AF38,
    ('PRESENTATION', 'AG38'): calc_PRESENTATION_AG38,
    ('PRESENTATION', 'H26'): calc_PRESENTATION_H26,
    ('PRESENTATION', 'O26'): calc_PRESENTATION_O26,
    ('PRESENTATION', 'D33'): calc_PRESENTATION_D33,
    ('PRESENTATION', 'D36'): calc_PRESENTATION_D36,
    ('PRESENTATION', 'F32'): calc_PRESENTATION_F32,
    ('PRESENTATION', 'M32'): calc_PRESENTATION_M32,
    ('PRESENTATION', 'R33'): calc_PRESENTATION_R33,
    ('PRESENTATION', 'R36'): calc_PRESENTATION_R36,
    ('PRESENTATION', 'T32'): calc_PRESENTATION_T32,
    ('PRESENTATION', 'V26'): calc_PRESENTATION_V26,
    ('PRESENTATION', 'Y33'): calc_PRESENTATION_Y33,
    ('PRESENTATION', 'Y36'): calc_PRESENTATION_Y36,
    ('PRESENTATION', 'AA32'): calc_PRESENTATION_AA32,
    ('PRESENTATION', 'AC26'): calc_PRESENTATION_AC26,
    ('PRESENTATION', 'K40'): calc_PRESENTATION_K40,
    ('Segment Revenue Model', 'T24'): calc_Segment_Revenue_Model_T24,
    ('Segment Revenue Model', 'W8'): calc_Segment_Revenue_Model_W8,
    ('Segment Revenue Model', 'T25'): calc_Segment_Revenue_Model_T25,
    ('Segment Revenue Model', 'W9'): calc_Segment_Revenue_Model_W9,
    ('Segment Revenue Model', 'T26'): calc_Segment_Revenue_Model_T26,
    ('Segment Revenue Model', 'W10'): calc_Segment_Revenue_Model_W10,
    ('Segment Revenue Model', 'T27'): calc_Segment_Revenue_Model_T27,
    ('Segment Revenue Model', 'W11'): calc_Segment_Revenue_Model_W11,
    ('Segment Revenue Model', 'X12'): calc_Segment_Revenue_Model_X12,
    ('Segment Revenue Model', 'T29'): calc_Segment_Revenue_Model_T29,
    ('Segment Revenue Model', 'W13'): calc_Segment_Revenue_Model_W13,
    ('Segment Revenue Model', 'K37'): calc_Segment_Revenue_Model_K37,
    ('Segment Revenue Model', 'K38'): calc_Segment_Revenue_Model_K38,
    ('Segment Revenue Model', 'K39'): calc_Segment_Revenue_Model_K39,
    ('Segment Revenue Model', 'K40'): calc_Segment_Revenue_Model_K40,
    ('Segment Revenue Model', 'K41'): calc_Segment_Revenue_Model_K41,
    ('Segment Revenue Model', 'K42'): calc_Segment_Revenue_Model_K42,
    ('Segment Revenue Model', 'K43'): calc_Segment_Revenue_Model_K43,
    ('Segment Revenue Model', 'K44'): calc_Segment_Revenue_Model_K44,
    ('Segment Revenue Model', 'K45'): calc_Segment_Revenue_Model_K45,
    ('Segment Revenue Model', 'K46'): calc_Segment_Revenue_Model_K46,
    ('Segment Revenue Model', 'T15'): calc_Segment_Revenue_Model_T15,
    ('Segment Revenue Model', 'T30'): calc_Segment_Revenue_Model_T30,
    ('Segment Revenue Model', 'W14'): calc_Segment_Revenue_Model_W14,
    ('Segment Revenue Model', 'V15'): calc_Segment_Revenue_Model_V15,
    ('Segment Revenue Model', 'P33'): calc_Segment_Revenue_Model_P33,
    ('Segment Revenue Model', 'P37'): calc_Segment_Revenue_Model_P37,
    ('Segment Revenue Model', 'P38'): calc_Segment_Revenue_Model_P38,
    ('Segment Revenue Model', 'P39'): calc_Segment_Revenue_Model_P39,
    ('Segment Revenue Model', 'P40'): calc_Segment_Revenue_Model_P40,
    ('Segment Revenue Model', 'P41'): calc_Segment_Revenue_Model_P41,
    ('Segment Revenue Model', 'P42'): calc_Segment_Revenue_Model_P42,
    ('Segment Revenue Model', 'P43'): calc_Segment_Revenue_Model_P43,
    ('Segment Revenue Model', 'P44'): calc_Segment_Revenue_Model_P44,
    ('Segment Revenue Model', 'P45'): calc_Segment_Revenue_Model_P45,
    ('Segment Revenue Model', 'P46'): calc_Segment_Revenue_Model_P46,
    ('Segment Revenue Model', 'K33'): calc_Segment_Revenue_Model_K33,
    ('Segment Revenue Model', 'X16'): calc_Segment_Revenue_Model_X16,
    ('INCOME STATEMENT', 'H7'): calc_INCOME_STATEMENT_H7,
    ('INCOME STATEMENT', 'I7'): calc_INCOME_STATEMENT_I7,
    ('INCOME STATEMENT', 'J7'): calc_INCOME_STATEMENT_J7,
    ('INCOME STATEMENT', 'K7'): calc_INCOME_STATEMENT_K7,
    ('INCOME STATEMENT', 'L7'): calc_INCOME_STATEMENT_L7,
    ('INCOME STATEMENT', 'M7'): calc_INCOME_STATEMENT_M7,
    ('INCOME STATEMENT', 'N7'): calc_INCOME_STATEMENT_N7,
    ('INCOME STATEMENT', 'O8'): calc_INCOME_STATEMENT_O8,
    ('INCOME STATEMENT', 'V8'): calc_INCOME_STATEMENT_V8,
    ('INCOME STATEMENT', 'O14'): calc_INCOME_STATEMENT_O14,
    ('INCOME STATEMENT', 'O20'): calc_INCOME_STATEMENT_O20,
    ('Ratio Analysis', 'D10'): calc_Ratio_Analysis_D10,
    ('Ratio Analysis', 'D21'): calc_Ratio_Analysis_D21,
    ('Ratio Analysis', 'E21'): calc_Ratio_Analysis_E21,
    ('Ratio Analysis', 'D23'): calc_Ratio_Analysis_D23,
    ('Ratio Analysis', 'D24'): calc_Ratio_Analysis_D24,
    ('INCOME STATEMENT', 'V7'): calc_INCOME_STATEMENT_V7,
    ('COMPANY OVERVIEW', 'F28'): calc_COMPANY_OVERVIEW_F28,
    ('INCOME STATEMENT', 'W7'): calc_INCOME_STATEMENT_W7,
    ('INCOME STATEMENT', 'X7'): calc_INCOME_STATEMENT_X7,
    ('INCOME STATEMENT', 'Y7'): calc_INCOME_STATEMENT_Y7,
    ('INCOME STATEMENT', 'Z7'): calc_INCOME_STATEMENT_Z7,
    ('INCOME STATEMENT', 'AA7'): calc_INCOME_STATEMENT_AA7,
    ('INCOME STATEMENT', 'AB7'): calc_INCOME_STATEMENT_AB7,
    ('INCOME STATEMENT', 'AC8'): calc_INCOME_STATEMENT_AC8,
    ('INCOME STATEMENT', 'AD8'): calc_INCOME_STATEMENT_AD8,
    ('Valuation', 'C19'): calc_Valuation_C19,
    ('Ratio Analysis', 'F10'): calc_Ratio_Analysis_F10,
    ('Ratio Analysis', 'F21'): calc_Ratio_Analysis_F21,
    ('Ratio Analysis', 'F23'): calc_Ratio_Analysis_F23,
    ('Ratio Analysis', 'F24'): calc_Ratio_Analysis_F24,
    ('INCOME STATEMENT', 'B30'): calc_INCOME_STATEMENT_B30,
    ('INCOME STATEMENT', 'C30'): calc_INCOME_STATEMENT_C30,
    ('INCOME STATEMENT', 'E30'): calc_INCOME_STATEMENT_E30,
    ('INCOME STATEMENT', 'G30'): calc_INCOME_STATEMENT_G30,
    ('INCOME STATEMENT', 'I30'): calc_INCOME_STATEMENT_I30,
    ('INCOME STATEMENT', 'J30'): calc_INCOME_STATEMENT_J30,
    ('INCOME STATEMENT', 'L30'): calc_INCOME_STATEMENT_L30,
    ('INCOME STATEMENT', 'N30'): calc_INCOME_STATEMENT_N30,
    ('INCOME STATEMENT', 'P30'): calc_INCOME_STATEMENT_P30,
    ('INCOME STATEMENT', 'Q30'): calc_INCOME_STATEMENT_Q30,
    ('INCOME STATEMENT', 'S30'): calc_INCOME_STATEMENT_S30,
    ('INCOME STATEMENT', 'U30'): calc_INCOME_STATEMENT_U30,
    ('INCOME STATEMENT', 'W30'): calc_INCOME_STATEMENT_W30,
    ('INCOME STATEMENT', 'X30'): calc_INCOME_STATEMENT_X30,
    ('INCOME STATEMENT', 'Z30'): calc_INCOME_STATEMENT_Z30,
    ('INCOME STATEMENT', 'AB30'): calc_INCOME_STATEMENT_AB30,
    ('INCOME STATEMENT', 'O12'): calc_INCOME_STATEMENT_O12,
    ('INCOME STATEMENT', 'AC12'): calc_INCOME_STATEMENT_AC12,
    ('INCOME STATEMENT', 'AC14'): calc_INCOME_STATEMENT_AC14,
    ('INCOME STATEMENT', 'H10'): calc_INCOME_STATEMENT_H10,
    ('INCOME STATEMENT', 'H16'): calc_INCOME_STATEMENT_H16,
    ('INCOME STATEMENT', 'O16'): calc_INCOME_STATEMENT_O16,
    ('INCOME STATEMENT', 'V16'): calc_INCOME_STATEMENT_V16,
    ('INCOME STATEMENT', 'AC16'): calc_INCOME_STATEMENT_AC16,
    ('INCOME STATEMENT', 'O10'): calc_INCOME_STATEMENT_O10,
    ('INCOME STATEMENT', 'O18'): calc_INCOME_STATEMENT_O18,
    ('INCOME STATEMENT', 'AC18'): calc_INCOME_STATEMENT_AC18,
    ('INCOME STATEMENT', 'K23'): calc_INCOME_STATEMENT_K23,
    ('INCOME STATEMENT', 'D25'): calc_INCOME_STATEMENT_D25,
    ('INCOME STATEMENT', 'D28'): calc_INCOME_STATEMENT_D28,
    ('INCOME STATEMENT', 'F21'): calc_INCOME_STATEMENT_F21,
    ('INCOME STATEMENT', 'H20'): calc_INCOME_STATEMENT_H20,
    ('INCOME STATEMENT', 'M21'): calc_INCOME_STATEMENT_M21,
    ('INCOME STATEMENT', 'R23'): calc_INCOME_STATEMENT_R23,
    ('INCOME STATEMENT', 'R25'): calc_INCOME_STATEMENT_R25,
    ('INCOME STATEMENT', 'R28'): calc_INCOME_STATEMENT_R28,
    ('INCOME STATEMENT', 'T21'): calc_INCOME_STATEMENT_T21,
    ('INCOME STATEMENT', 'V10'): calc_INCOME_STATEMENT_V10,
    ('INCOME STATEMENT', 'V20'): calc_INCOME_STATEMENT_V20,
    ('INCOME STATEMENT', 'Y23'): calc_INCOME_STATEMENT_Y23,
    ('INCOME STATEMENT', 'Y25'): calc_INCOME_STATEMENT_Y25,
    ('INCOME STATEMENT', 'Y28'): calc_INCOME_STATEMENT_Y28,
    ('INCOME STATEMENT', 'AA21'): calc_INCOME_STATEMENT_AA21,
    ('INCOME STATEMENT', 'AC10'): calc_INCOME_STATEMENT_AC10,
    ('INCOME STATEMENT', 'AC20'): calc_INCOME_STATEMENT_AC20,
    ('COMPANY OVERVIEW', 'G22'): calc_COMPANY_OVERVIEW_G22,
    ('INCOME STATEMENT', 'AD25'): calc_INCOME_STATEMENT_AD25,
    ('INCOME STATEMENT', 'AD28'): calc_INCOME_STATEMENT_AD28,
    ('Valuation', 'D20'): calc_Valuation_D20,
    ('Ratio Analysis', 'G13'): calc_Ratio_Analysis_G13,
    ('Ratio Analysis', 'G19'): calc_Ratio_Analysis_G19,
    ('Ratio Analysis', 'G22'): calc_Ratio_Analysis_G22,
    ('COMPANY OVERVIEW', 'H22'): calc_COMPANY_OVERVIEW_H22,
    ('INCOME STATEMENT', 'AE23'): calc_INCOME_STATEMENT_AE23,
    ('INCOME STATEMENT', 'AE25'): calc_INCOME_STATEMENT_AE25,
    ('INCOME STATEMENT', 'AE28'): calc_INCOME_STATEMENT_AE28,
    ('Valuation', 'E20'): calc_Valuation_E20,
    ('Ratio Analysis', 'H13'): calc_Ratio_Analysis_H13,
    ('Ratio Analysis', 'H19'): calc_Ratio_Analysis_H19,
    ('Ratio Analysis', 'H22'): calc_Ratio_Analysis_H22,
    ('COMPANY OVERVIEW', 'I22'): calc_COMPANY_OVERVIEW_I22,
    ('INCOME STATEMENT', 'AF23'): calc_INCOME_STATEMENT_AF23,
    ('INCOME STATEMENT', 'AF25'): calc_INCOME_STATEMENT_AF25,
    ('INCOME STATEMENT', 'AF28'): calc_INCOME_STATEMENT_AF28,
    ('Valuation', 'F20'): calc_Valuation_F20,
    ('Ratio Analysis', 'I13'): calc_Ratio_Analysis_I13,
    ('Ratio Analysis', 'I19'): calc_Ratio_Analysis_I19,
    ('Ratio Analysis', 'I22'): calc_Ratio_Analysis_I22,
    ('COMPANY OVERVIEW', 'J22'): calc_COMPANY_OVERVIEW_J22,
    ('INCOME STATEMENT', 'AG23'): calc_INCOME_STATEMENT_AG23,
    ('INCOME STATEMENT', 'AG25'): calc_INCOME_STATEMENT_AG25,
    ('INCOME STATEMENT', 'AG28'): calc_INCOME_STATEMENT_AG28,
    ('Valuation', 'G20'): calc_Valuation_G20,
    ('Ratio Analysis', 'J13'): calc_Ratio_Analysis_J13,
    ('Ratio Analysis', 'J19'): calc_Ratio_Analysis_J19,
    ('Ratio Analysis', 'J22'): calc_Ratio_Analysis_J22,
    ('Ratio Analysis', 'C15'): calc_Ratio_Analysis_C15,
    ('Ratio Analysis', 'D15'): calc_Ratio_Analysis_D15,
    ('Ratio Analysis', 'E15'): calc_Ratio_Analysis_E15,
    ('Ratio Analysis', 'F15'): calc_Ratio_Analysis_F15,
    ('CASH FOW STATEMENT', 'B11'): calc_CASH_FOW_STATEMENT_B11,
    ('CASH FOW STATEMENT', 'C11'): calc_CASH_FOW_STATEMENT_C11,
    ('INCOME STATEMENT', 'K32'): calc_INCOME_STATEMENT_K32,
    ('CASH FOW STATEMENT', 'D11'): calc_CASH_FOW_STATEMENT_D11,
    ('Valuation', 'B24'): calc_Valuation_B24,
    ('Valuation', 'C22'): calc_Valuation_C22,
    ('CASH FOW STATEMENT', 'B9'): calc_CASH_FOW_STATEMENT_B9,
    ('CASH FOW STATEMENT', 'E9'): calc_CASH_FOW_STATEMENT_E9,
    ('INCOME STATEMENT', 'F47'): calc_INCOME_STATEMENT_F47,
    ('INCOME STATEMENT', 'E48'): calc_INCOME_STATEMENT_E48,
    ('INCOME STATEMENT', 'W47'): calc_INCOME_STATEMENT_W47,
    ('INCOME STATEMENT', 'V48'): calc_INCOME_STATEMENT_V48,
    ('CASH FOW STATEMENT', 'I14'): calc_CASH_FOW_STATEMENT_I14,
    ('Valuation', 'F46'): calc_Valuation_F46,
    ('Valuation', 'I25'): calc_Valuation_I25,
    ('Valuation', 'G13'): calc_Valuation_G13,
    ('Valuation', 'G55'): calc_Valuation_G55,
    ('Valuation', 'B25'): calc_Valuation_B25,
    ('Valuation', 'C25'): calc_Valuation_C25,
    ('Valuation', 'G37'): calc_Valuation_G37,
    ('Valuation', 'G45'): calc_Valuation_G45,
    ('PRESENTATION', 'B40'): calc_PRESENTATION_B40,
    ('PRESENTATION', 'C40'): calc_PRESENTATION_C40,
    ('PRESENTATION', 'E40'): calc_PRESENTATION_E40,
    ('PRESENTATION', 'G40'): calc_PRESENTATION_G40,
    ('PRESENTATION', 'I40'): calc_PRESENTATION_I40,
    ('PRESENTATION', 'J40'): calc_PRESENTATION_J40,
    ('PRESENTATION', 'L40'): calc_PRESENTATION_L40,
    ('PRESENTATION', 'N40'): calc_PRESENTATION_N40,
    ('PRESENTATION', 'P40'): calc_PRESENTATION_P40,
    ('PRESENTATION', 'Q40'): calc_PRESENTATION_Q40,
    ('PRESENTATION', 'S40'): calc_PRESENTATION_S40,
    ('PRESENTATION', 'U40'): calc_PRESENTATION_U40,
    ('PRESENTATION', 'W40'): calc_PRESENTATION_W40,
    ('PRESENTATION', 'X40'): calc_PRESENTATION_X40,
    ('PRESENTATION', 'Z40'): calc_PRESENTATION_Z40,
    ('PRESENTATION', 'AB40'): calc_PRESENTATION_AB40,
    ('PRESENTATION', 'AD40'): calc_PRESENTATION_AD40,
    ('PRESENTATION', 'AE40'): calc_PRESENTATION_AE40,
    ('PRESENTATION', 'AF40'): calc_PRESENTATION_AF40,
    ('PRESENTATION', 'AG40'): calc_PRESENTATION_AG40,
    ('PRESENTATION', 'H32'): calc_PRESENTATION_H32,
    ('PRESENTATION', 'O32'): calc_PRESENTATION_O32,
    ('PRESENTATION', 'D38'): calc_PRESENTATION_D38,
    ('PRESENTATION', 'F33'): calc_PRESENTATION_F33,
    ('PRESENTATION', 'F36'): calc_PRESENTATION_F36,
    ('PRESENTATION', 'M33'): calc_PRESENTATION_M33,
    ('PRESENTATION', 'M36'): calc_PRESENTATION_M36,
    ('PRESENTATION', 'R38'): calc_PRESENTATION_R38,
    ('PRESENTATION', 'T33'): calc_PRESENTATION_T33,
    ('PRESENTATION', 'T36'): calc_PRESENTATION_T36,
    ('PRESENTATION', 'V32'): calc_PRESENTATION_V32,
    ('PRESENTATION', 'Y38'): calc_PRESENTATION_Y38,
    ('PRESENTATION', 'AA33'): calc_PRESENTATION_AA33,
    ('PRESENTATION', 'AA36'): calc_PRESENTATION_AA36,
    ('PRESENTATION', 'AC32'): calc_PRESENTATION_AC32,
    ('PRESENTATION', 'K44'): calc_PRESENTATION_K44,
    ('Segment Revenue Model', 'X8'): calc_Segment_Revenue_Model_X8,
    ('Segment Revenue Model', 'X9'): calc_Segment_Revenue_Model_X9,
    ('Segment Revenue Model', 'X10'): calc_Segment_Revenue_Model_X10,
    ('Segment Revenue Model', 'X11'): calc_Segment_Revenue_Model_X11,
    ('Segment Revenue Model', 'Y12'): calc_Segment_Revenue_Model_Y12,
    ('Segment Revenue Model', 'X13'): calc_Segment_Revenue_Model_X13,
    ('Segment Revenue Model', 'U15'): calc_Segment_Revenue_Model_U15,
    ('Segment Revenue Model', 'T17'): calc_Segment_Revenue_Model_T17,
    ('Segment Revenue Model', 'T31'): calc_Segment_Revenue_Model_T31,
    ('Segment Revenue Model', 'X14'): calc_Segment_Revenue_Model_X14,
    ('Segment Revenue Model', 'W15'): calc_Segment_Revenue_Model_W15,
    ('Segment Revenue Model', 'V17'): calc_Segment_Revenue_Model_V17,
    ('Segment Revenue Model', 'Y16'): calc_Segment_Revenue_Model_Y16,
    ('INCOME STATEMENT', 'O7'): calc_INCOME_STATEMENT_O7,
    ('INCOME STATEMENT', 'AC7'): calc_INCOME_STATEMENT_AC7,
    ('Valuation', 'C37'): calc_Valuation_C37,
    ('Valuation', 'D37'): calc_Valuation_D37,
    ('Valuation', 'C45'): calc_Valuation_C45,
    ('Valuation', 'C46'): calc_Valuation_C46,
    ('INCOME STATEMENT', 'B32'): calc_INCOME_STATEMENT_B32,
    ('INCOME STATEMENT', 'C32'): calc_INCOME_STATEMENT_C32,
    ('INCOME STATEMENT', 'E32'): calc_INCOME_STATEMENT_E32,
    ('INCOME STATEMENT', 'G32'): calc_INCOME_STATEMENT_G32,
    ('INCOME STATEMENT', 'I32'): calc_INCOME_STATEMENT_I32,
    ('INCOME STATEMENT', 'J32'): calc_INCOME_STATEMENT_J32,
    ('INCOME STATEMENT', 'L32'): calc_INCOME_STATEMENT_L32,
    ('INCOME STATEMENT', 'N32'): calc_INCOME_STATEMENT_N32,
    ('INCOME STATEMENT', 'P32'): calc_INCOME_STATEMENT_P32,
    ('INCOME STATEMENT', 'Q32'): calc_INCOME_STATEMENT_Q32,
    ('INCOME STATEMENT', 'S32'): calc_INCOME_STATEMENT_S32,
    ('INCOME STATEMENT', 'U32'): calc_INCOME_STATEMENT_U32,
    ('INCOME STATEMENT', 'W32'): calc_INCOME_STATEMENT_W32,
    ('INCOME STATEMENT', 'X32'): calc_INCOME_STATEMENT_X32,
    ('INCOME STATEMENT', 'Z32'): calc_INCOME_STATEMENT_Z32,
    ('INCOME STATEMENT', 'AB32'): calc_INCOME_STATEMENT_AB32,
    ('PRESENTATION', 'C64'): calc_PRESENTATION_C64,
    ('INCOME STATEMENT', 'H21'): calc_INCOME_STATEMENT_H21,
    ('PRESENTATION', 'D64'): calc_PRESENTATION_D64,
    ('INCOME STATEMENT', 'O21'): calc_INCOME_STATEMENT_O21,
    ('INCOME STATEMENT', 'D30'): calc_INCOME_STATEMENT_D30,
    ('INCOME STATEMENT', 'F25'): calc_INCOME_STATEMENT_F25,
    ('INCOME STATEMENT', 'F28'): calc_INCOME_STATEMENT_F28,
    ('INCOME STATEMENT', 'M23'): calc_INCOME_STATEMENT_M23,
    ('INCOME STATEMENT', 'M25'): calc_INCOME_STATEMENT_M25,
    ('INCOME STATEMENT', 'M28'): calc_INCOME_STATEMENT_M28,
    ('INCOME STATEMENT', 'R30'): calc_INCOME_STATEMENT_R30,
    ('INCOME STATEMENT', 'T23'): calc_INCOME_STATEMENT_T23,
    ('INCOME STATEMENT', 'T25'): calc_INCOME_STATEMENT_T25,
    ('INCOME STATEMENT', 'T28'): calc_INCOME_STATEMENT_T28,
    ('PRESENTATION', 'E64'): calc_PRESENTATION_E64,
    ('INCOME STATEMENT', 'V21'): calc_INCOME_STATEMENT_V21,
    ('INCOME STATEMENT', 'Y30'): calc_INCOME_STATEMENT_Y30,
    ('INCOME STATEMENT', 'AA23'): calc_INCOME_STATEMENT_AA23,
    ('INCOME STATEMENT', 'AA25'): calc_INCOME_STATEMENT_AA25,
    ('INCOME STATEMENT', 'AA28'): calc_INCOME_STATEMENT_AA28,
    ('PRESENTATION', 'F64'): calc_PRESENTATION_F64,
    ('INCOME STATEMENT', 'AC21'): calc_INCOME_STATEMENT_AC21,
    ('INCOME STATEMENT', 'AD30'): calc_INCOME_STATEMENT_AD30,
    ('Valuation', 'D21'): calc_Valuation_D21,
    ('Valuation', 'D43'): calc_Valuation_D43,
    ('PRESENTATION', 'G63'): calc_PRESENTATION_G63,
    ('COMPANY OVERVIEW', 'G24'): calc_COMPANY_OVERVIEW_G24,
    ('PRESENTATION', 'G66'): calc_PRESENTATION_G66,
    ('INCOME STATEMENT', 'AE30'): calc_INCOME_STATEMENT_AE30,
    ('Valuation', 'E21'): calc_Valuation_E21,
    ('Valuation', 'E38'): calc_Valuation_E38,
    ('Valuation', 'E43'): calc_Valuation_E43,
    ('PRESENTATION', 'H63'): calc_PRESENTATION_H63,
    ('COMPANY OVERVIEW', 'H24'): calc_COMPANY_OVERVIEW_H24,
    ('PRESENTATION', 'H66'): calc_PRESENTATION_H66,
    ('INCOME STATEMENT', 'AF30'): calc_INCOME_STATEMENT_AF30,
    ('Valuation', 'F21'): calc_Valuation_F21,
    ('Valuation', 'F38'): calc_Valuation_F38,
    ('PRESENTATION', 'I63'): calc_PRESENTATION_I63,
    ('COMPANY OVERVIEW', 'I24'): calc_COMPANY_OVERVIEW_I24,
    ('PRESENTATION', 'I66'): calc_PRESENTATION_I66,
    ('INCOME STATEMENT', 'AG30'): calc_INCOME_STATEMENT_AG30,
    ('Valuation', 'G21'): calc_Valuation_G21,
    ('Valuation', 'G38'): calc_Valuation_G38,
    ('PRESENTATION', 'J63'): calc_PRESENTATION_J63,
    ('COMPANY OVERVIEW', 'J24'): calc_COMPANY_OVERVIEW_J24,
    ('PRESENTATION', 'J66'): calc_PRESENTATION_J66,
    ('INCOME STATEMENT', 'K36'): calc_INCOME_STATEMENT_K36,
    ('INCOME STATEMENT', 'K37'): calc_INCOME_STATEMENT_K37,
    ('INCOME STATEMENT', 'K38'): calc_INCOME_STATEMENT_K38,
    ('Valuation', 'C40'): calc_Valuation_C40,
    ('Valuation', 'B45'): calc_Valuation_B45,
    ('INCOME STATEMENT', 'G47'): calc_INCOME_STATEMENT_G47,
    ('INCOME STATEMENT', 'F48'): calc_INCOME_STATEMENT_F48,
    ('INCOME STATEMENT', 'X47'): calc_INCOME_STATEMENT_X47,
    ('INCOME STATEMENT', 'W48'): calc_INCOME_STATEMENT_W48,
    ('Valuation', 'B30'): calc_Valuation_B30,
    ('CASH FOW STATEMENT', 'I23'): calc_CASH_FOW_STATEMENT_I23,
    ('Valuation', 'G46'): calc_Valuation_G46,
    ('Valuation', 'J25'): calc_Valuation_J25,
    ('COMPANY OVERVIEW', 'E25'): calc_COMPANY_OVERVIEW_E25,
    ('COMPANY OVERVIEW', 'F25'): calc_COMPANY_OVERVIEW_F25,
    ('COMPANY OVERVIEW', 'G25'): calc_COMPANY_OVERVIEW_G25,
    ('COMPANY OVERVIEW', 'H25'): calc_COMPANY_OVERVIEW_H25,
    ('COMPANY OVERVIEW', 'I25'): calc_COMPANY_OVERVIEW_I25,
    ('COMPANY OVERVIEW', 'J25'): calc_COMPANY_OVERVIEW_J25,
    ('COMPANY OVERVIEW', 'G26'): calc_COMPANY_OVERVIEW_G26,
    ('COMPANY OVERVIEW', 'H26'): calc_COMPANY_OVERVIEW_H26,
    ('COMPANY OVERVIEW', 'I26'): calc_COMPANY_OVERVIEW_I26,
    ('COMPANY OVERVIEW', 'J26'): calc_COMPANY_OVERVIEW_J26,
    ('Valuation', 'K6'): calc_Valuation_K6,
    ('Valuation', 'K7'): calc_Valuation_K7,
    ('Valuation', 'H37'): calc_Valuation_H37,
    ('Valuation', 'H45'): calc_Valuation_H45,
    ('PRESENTATION', 'B44'): calc_PRESENTATION_B44,
    ('PRESENTATION', 'C44'): calc_PRESENTATION_C44,
    ('PRESENTATION', 'E44'): calc_PRESENTATION_E44,
    ('PRESENTATION', 'G44'): calc_PRESENTATION_G44,
    ('PRESENTATION', 'I44'): calc_PRESENTATION_I44,
    ('PRESENTATION', 'J44'): calc_PRESENTATION_J44,
    ('PRESENTATION', 'L44'): calc_PRESENTATION_L44,
    ('PRESENTATION', 'N44'): calc_PRESENTATION_N44,
    ('PRESENTATION', 'P44'): calc_PRESENTATION_P44,
    ('PRESENTATION', 'Q44'): calc_PRESENTATION_Q44,
    ('PRESENTATION', 'S44'): calc_PRESENTATION_S44,
    ('PRESENTATION', 'U44'): calc_PRESENTATION_U44,
    ('PRESENTATION', 'W44'): calc_PRESENTATION_W44,
    ('PRESENTATION', 'X44'): calc_PRESENTATION_X44,
    ('PRESENTATION', 'Z44'): calc_PRESENTATION_Z44,
    ('PRESENTATION', 'AB44'): calc_PRESENTATION_AB44,
    ('PRESENTATION', 'AD41'): calc_PRESENTATION_AD41,
    ('PRESENTATION', 'AE41'): calc_PRESENTATION_AE41,
    ('PRESENTATION', 'AF41'): calc_PRESENTATION_AF41,
    ('PRESENTATION', 'AG41'): calc_PRESENTATION_AG41,
    ('PRESENTATION', 'H33'): calc_PRESENTATION_H33,
    ('PRESENTATION', 'H36'): calc_PRESENTATION_H36,
    ('PRESENTATION', 'O33'): calc_PRESENTATION_O33,
    ('PRESENTATION', 'O36'): calc_PRESENTATION_O36,
    ('PRESENTATION', 'D40'): calc_PRESENTATION_D40,
    ('PRESENTATION', 'F38'): calc_PRESENTATION_F38,
    ('PRESENTATION', 'M38'): calc_PRESENTATION_M38,
    ('PRESENTATION', 'R40'): calc_PRESENTATION_R40,
    ('PRESENTATION', 'T38'): calc_PRESENTATION_T38,
    ('PRESENTATION', 'V33'): calc_PRESENTATION_V33,
    ('PRESENTATION', 'V36'): calc_PRESENTATION_V36,
    ('PRESENTATION', 'Y40'): calc_PRESENTATION_Y40,
    ('PRESENTATION', 'AA38'): calc_PRESENTATION_AA38,
    ('PRESENTATION', 'AC33'): calc_PRESENTATION_AC33,
    ('PRESENTATION', 'AC36'): calc_PRESENTATION_AC36,
    ('PRESENTATION', 'K45'): calc_PRESENTATION_K45,
    ('PRESENTATION', 'K47'): calc_PRESENTATION_K47,
    ('Segment Revenue Model', 'Y8'): calc_Segment_Revenue_Model_Y8,
    ('Segment Revenue Model', 'Y9'): calc_Segment_Revenue_Model_Y9,
    ('Segment Revenue Model', 'Y10'): calc_Segment_Revenue_Model_Y10,
    ('Segment Revenue Model', 'Y11'): calc_Segment_Revenue_Model_Y11,
    ('Segment Revenue Model', 'Y13'): calc_Segment_Revenue_Model_Y13,
    ('Segment Revenue Model', 'U17'): calc_Segment_Revenue_Model_U17,
    ('Segment Revenue Model', 'T33'): calc_Segment_Revenue_Model_T33,
    ('Segment Revenue Model', 'T37'): calc_Segment_Revenue_Model_T37,
    ('Segment Revenue Model', 'T38'): calc_Segment_Revenue_Model_T38,
    ('Segment Revenue Model', 'T39'): calc_Segment_Revenue_Model_T39,
    ('Segment Revenue Model', 'T40'): calc_Segment_Revenue_Model_T40,
    ('Segment Revenue Model', 'T41'): calc_Segment_Revenue_Model_T41,
    ('Segment Revenue Model', 'T42'): calc_Segment_Revenue_Model_T42,
    ('Segment Revenue Model', 'T43'): calc_Segment_Revenue_Model_T43,
    ('Segment Revenue Model', 'T44'): calc_Segment_Revenue_Model_T44,
    ('Segment Revenue Model', 'T45'): calc_Segment_Revenue_Model_T45,
    ('Segment Revenue Model', 'T46'): calc_Segment_Revenue_Model_T46,
    ('Segment Revenue Model', 'Y14'): calc_Segment_Revenue_Model_Y14,
    ('Segment Revenue Model', 'X15'): calc_Segment_Revenue_Model_X15,
    ('Segment Revenue Model', 'W17'): calc_Segment_Revenue_Model_W17,
    ('Segment Revenue Model', 'V37'): calc_Segment_Revenue_Model_V37,
    ('Segment Revenue Model', 'V38'): calc_Segment_Revenue_Model_V38,
    ('Segment Revenue Model', 'V39'): calc_Segment_Revenue_Model_V39,
    ('Segment Revenue Model', 'V40'): calc_Segment_Revenue_Model_V40,
    ('Segment Revenue Model', 'V41'): calc_Segment_Revenue_Model_V41,
    ('Segment Revenue Model', 'V42'): calc_Segment_Revenue_Model_V42,
    ('Segment Revenue Model', 'V43'): calc_Segment_Revenue_Model_V43,
    ('Segment Revenue Model', 'V44'): calc_Segment_Revenue_Model_V44,
    ('Segment Revenue Model', 'V45'): calc_Segment_Revenue_Model_V45,
    ('Segment Revenue Model', 'V46'): calc_Segment_Revenue_Model_V46,
    ('INCOME STATEMENT', 'B36'): calc_INCOME_STATEMENT_B36,
    ('INCOME STATEMENT', 'B37'): calc_INCOME_STATEMENT_B37,
    ('INCOME STATEMENT', 'B38'): calc_INCOME_STATEMENT_B38,
    ('INCOME STATEMENT', 'C36'): calc_INCOME_STATEMENT_C36,
    ('INCOME STATEMENT', 'C37'): calc_INCOME_STATEMENT_C37,
    ('INCOME STATEMENT', 'C38'): calc_INCOME_STATEMENT_C38,
    ('INCOME STATEMENT', 'E36'): calc_INCOME_STATEMENT_E36,
    ('INCOME STATEMENT', 'E37'): calc_INCOME_STATEMENT_E37,
    ('INCOME STATEMENT', 'E38'): calc_INCOME_STATEMENT_E38,
    ('INCOME STATEMENT', 'G36'): calc_INCOME_STATEMENT_G36,
    ('INCOME STATEMENT', 'G37'): calc_INCOME_STATEMENT_G37,
    ('INCOME STATEMENT', 'G38'): calc_INCOME_STATEMENT_G38,
    ('INCOME STATEMENT', 'I36'): calc_INCOME_STATEMENT_I36,
    ('INCOME STATEMENT', 'I37'): calc_INCOME_STATEMENT_I37,
    ('INCOME STATEMENT', 'I38'): calc_INCOME_STATEMENT_I38,
    ('INCOME STATEMENT', 'J36'): calc_INCOME_STATEMENT_J36,
    ('INCOME STATEMENT', 'J37'): calc_INCOME_STATEMENT_J37,
    ('INCOME STATEMENT', 'J38'): calc_INCOME_STATEMENT_J38,
    ('INCOME STATEMENT', 'L36'): calc_INCOME_STATEMENT_L36,
    ('INCOME STATEMENT', 'L37'): calc_INCOME_STATEMENT_L37,
    ('INCOME STATEMENT', 'L38'): calc_INCOME_STATEMENT_L38,
    ('INCOME STATEMENT', 'N36'): calc_INCOME_STATEMENT_N36,
    ('INCOME STATEMENT', 'N37'): calc_INCOME_STATEMENT_N37,
    ('INCOME STATEMENT', 'N38'): calc_INCOME_STATEMENT_N38,
    ('INCOME STATEMENT', 'P36'): calc_INCOME_STATEMENT_P36,
    ('INCOME STATEMENT', 'P37'): calc_INCOME_STATEMENT_P37,
    ('INCOME STATEMENT', 'P38'): calc_INCOME_STATEMENT_P38,
    ('INCOME STATEMENT', 'Q36'): calc_INCOME_STATEMENT_Q36,
    ('INCOME STATEMENT', 'Q37'): calc_INCOME_STATEMENT_Q37,
    ('INCOME STATEMENT', 'Q38'): calc_INCOME_STATEMENT_Q38,
    ('INCOME STATEMENT', 'S36'): calc_INCOME_STATEMENT_S36,
    ('INCOME STATEMENT', 'S37'): calc_INCOME_STATEMENT_S37,
    ('INCOME STATEMENT', 'S38'): calc_INCOME_STATEMENT_S38,
    ('INCOME STATEMENT', 'U36'): calc_INCOME_STATEMENT_U36,
    ('INCOME STATEMENT', 'U37'): calc_INCOME_STATEMENT_U37,
    ('INCOME STATEMENT', 'U38'): calc_INCOME_STATEMENT_U38,
    ('INCOME STATEMENT', 'W36'): calc_INCOME_STATEMENT_W36,
    ('INCOME STATEMENT', 'W37'): calc_INCOME_STATEMENT_W37,
    ('INCOME STATEMENT', 'W38'): calc_INCOME_STATEMENT_W38,
    ('INCOME STATEMENT', 'X36'): calc_INCOME_STATEMENT_X36,
    ('INCOME STATEMENT', 'X37'): calc_INCOME_STATEMENT_X37,
    ('INCOME STATEMENT', 'X38'): calc_INCOME_STATEMENT_X38,
    ('INCOME STATEMENT', 'Z36'): calc_INCOME_STATEMENT_Z36,
    ('INCOME STATEMENT', 'Z37'): calc_INCOME_STATEMENT_Z37,
    ('INCOME STATEMENT', 'Z38'): calc_INCOME_STATEMENT_Z38,
    ('INCOME STATEMENT', 'AB36'): calc_INCOME_STATEMENT_AB36,
    ('INCOME STATEMENT', 'AB37'): calc_INCOME_STATEMENT_AB37,
    ('INCOME STATEMENT', 'AB38'): calc_INCOME_STATEMENT_AB38,
    ('COMPANY OVERVIEW', 'C22'): calc_COMPANY_OVERVIEW_C22,
    ('INCOME STATEMENT', 'B22'): calc_INCOME_STATEMENT_B22,
    ('INCOME STATEMENT', 'C22'): calc_INCOME_STATEMENT_C22,
    ('INCOME STATEMENT', 'D22'): calc_INCOME_STATEMENT_D22,
    ('INCOME STATEMENT', 'E22'): calc_INCOME_STATEMENT_E22,
    ('INCOME STATEMENT', 'F22'): calc_INCOME_STATEMENT_F22,
    ('INCOME STATEMENT', 'G22'): calc_INCOME_STATEMENT_G22,
    ('INCOME STATEMENT', 'H25'): calc_INCOME_STATEMENT_H25,
    ('INCOME STATEMENT', 'H28'): calc_INCOME_STATEMENT_H28,
    ('Ratio Analysis', 'C13'): calc_Ratio_Analysis_C13,
    ('Ratio Analysis', 'C19'): calc_Ratio_Analysis_C19,
    ('COMPANY OVERVIEW', 'D22'): calc_COMPANY_OVERVIEW_D22,
    ('INCOME STATEMENT', 'I22'): calc_INCOME_STATEMENT_I22,
    ('INCOME STATEMENT', 'J22'): calc_INCOME_STATEMENT_J22,
    ('INCOME STATEMENT', 'K22'): calc_INCOME_STATEMENT_K22,
    ('INCOME STATEMENT', 'L22'): calc_INCOME_STATEMENT_L22,
    ('INCOME STATEMENT', 'M22'): calc_INCOME_STATEMENT_M22,
    ('INCOME STATEMENT', 'N22'): calc_INCOME_STATEMENT_N22,
    ('INCOME STATEMENT', 'O23'): calc_INCOME_STATEMENT_O23,
    ('INCOME STATEMENT', 'O25'): calc_INCOME_STATEMENT_O25,
    ('INCOME STATEMENT', 'O28'): calc_INCOME_STATEMENT_O28,
    ('Ratio Analysis', 'D13'): calc_Ratio_Analysis_D13,
    ('Ratio Analysis', 'D19'): calc_Ratio_Analysis_D19,
    ('Ratio Analysis', 'D22'): calc_Ratio_Analysis_D22,
    ('INCOME STATEMENT', 'D32'): calc_INCOME_STATEMENT_D32,
    ('INCOME STATEMENT', 'F30'): calc_INCOME_STATEMENT_F30,
    ('INCOME STATEMENT', 'M30'): calc_INCOME_STATEMENT_M30,
    ('INCOME STATEMENT', 'R32'): calc_INCOME_STATEMENT_R32,
    ('INCOME STATEMENT', 'T30'): calc_INCOME_STATEMENT_T30,
    ('COMPANY OVERVIEW', 'E22'): calc_COMPANY_OVERVIEW_E22,
    ('COMPANY OVERVIEW', 'E26'): calc_COMPANY_OVERVIEW_E26,
    ('INCOME STATEMENT', 'P22'): calc_INCOME_STATEMENT_P22,
    ('INCOME STATEMENT', 'Q22'): calc_INCOME_STATEMENT_Q22,
    ('INCOME STATEMENT', 'R22'): calc_INCOME_STATEMENT_R22,
    ('INCOME STATEMENT', 'S22'): calc_INCOME_STATEMENT_S22,
    ('INCOME STATEMENT', 'T22'): calc_INCOME_STATEMENT_T22,
    ('INCOME STATEMENT', 'U22'): calc_INCOME_STATEMENT_U22,
    ('INCOME STATEMENT', 'V23'): calc_INCOME_STATEMENT_V23,
    ('INCOME STATEMENT', 'V25'): calc_INCOME_STATEMENT_V25,
    ('INCOME STATEMENT', 'V28'): calc_INCOME_STATEMENT_V28,
    ('Valuation', 'B20'): calc_Valuation_B20,
    ('Valuation', 'B21'): calc_Valuation_B21,
    ('Ratio Analysis', 'E13'): calc_Ratio_Analysis_E13,
    ('Ratio Analysis', 'E19'): calc_Ratio_Analysis_E19,
    ('Ratio Analysis', 'E22'): calc_Ratio_Analysis_E22,
    ('INCOME STATEMENT', 'Y32'): calc_INCOME_STATEMENT_Y32,
    ('INCOME STATEMENT', 'AA30'): calc_INCOME_STATEMENT_AA30,
    ('COMPANY OVERVIEW', 'F22'): calc_COMPANY_OVERVIEW_F22,
    ('COMPANY OVERVIEW', 'F26'): calc_COMPANY_OVERVIEW_F26,
    ('INCOME STATEMENT', 'W22'): calc_INCOME_STATEMENT_W22,
    ('INCOME STATEMENT', 'X22'): calc_INCOME_STATEMENT_X22,
    ('INCOME STATEMENT', 'Y22'): calc_INCOME_STATEMENT_Y22,
    ('INCOME STATEMENT', 'Z22'): calc_INCOME_STATEMENT_Z22,
    ('INCOME STATEMENT', 'AA22'): calc_INCOME_STATEMENT_AA22,
    ('INCOME STATEMENT', 'AB22'): calc_INCOME_STATEMENT_AB22,
    ('INCOME STATEMENT', 'AC23'): calc_INCOME_STATEMENT_AC23,
    ('INCOME STATEMENT', 'AD23'): calc_INCOME_STATEMENT_AD23,
    ('INCOME STATEMENT', 'AC25'): calc_INCOME_STATEMENT_AC25,
    ('INCOME STATEMENT', 'AC28'): calc_INCOME_STATEMENT_AC28,
    ('Valuation', 'C20'): calc_Valuation_C20,
    ('Ratio Analysis', 'F13'): calc_Ratio_Analysis_F13,
    ('Ratio Analysis', 'F19'): calc_Ratio_Analysis_F19,
    ('Ratio Analysis', 'F22'): calc_Ratio_Analysis_F22,
    ('INCOME STATEMENT', 'AD32'): calc_INCOME_STATEMENT_AD32,
    ('Ratio Analysis', 'G26'): calc_Ratio_Analysis_G26,
    ('Valuation', 'D44'): calc_Valuation_D44,
    ('INCOME STATEMENT', 'AE32'): calc_INCOME_STATEMENT_AE32,
    ('Ratio Analysis', 'H26'): calc_Ratio_Analysis_H26,
    ('Valuation', 'E39'): calc_Valuation_E39,
    ('Valuation', 'E44'): calc_Valuation_E44,
    ('INCOME STATEMENT', 'AF32'): calc_INCOME_STATEMENT_AF32,
    ('Ratio Analysis', 'I26'): calc_Ratio_Analysis_I26,
    ('Valuation', 'F39'): calc_Valuation_F39,
    ('INCOME STATEMENT', 'AG32'): calc_INCOME_STATEMENT_AG32,
    ('Ratio Analysis', 'J26'): calc_Ratio_Analysis_J26,
    ('Valuation', 'G39'): calc_Valuation_G39,
    ('INCOME STATEMENT', 'K41'): calc_INCOME_STATEMENT_K41,
    ('INCOME STATEMENT', 'H47'): calc_INCOME_STATEMENT_H47,
    ('INCOME STATEMENT', 'G48'): calc_INCOME_STATEMENT_G48,
    ('INCOME STATEMENT', 'Y47'): calc_INCOME_STATEMENT_Y47,
    ('INCOME STATEMENT', 'X48'): calc_INCOME_STATEMENT_X48,
    ('CASH FOW STATEMENT', 'I24'): calc_CASH_FOW_STATEMENT_I24,
    ('Valuation', 'H46'): calc_Valuation_H46,
    ('Valuation', 'K25'): calc_Valuation_K25,
    ('Valuation', 'H19'): calc_Valuation_H19,
    ('Valuation', 'I37'): calc_Valuation_I37,
    ('Valuation', 'I45'): calc_Valuation_I45,
    ('PRESENTATION', 'B45'): calc_PRESENTATION_B45,
    ('PRESENTATION', 'C45'): calc_PRESENTATION_C45,
    ('PRESENTATION', 'C47'): calc_PRESENTATION_C47,
    ('PRESENTATION', 'E45'): calc_PRESENTATION_E45,
    ('PRESENTATION', 'E47'): calc_PRESENTATION_E47,
    ('PRESENTATION', 'G45'): calc_PRESENTATION_G45,
    ('PRESENTATION', 'G47'): calc_PRESENTATION_G47,
    ('PRESENTATION', 'I45'): calc_PRESENTATION_I45,
    ('PRESENTATION', 'I47'): calc_PRESENTATION_I47,
    ('PRESENTATION', 'J45'): calc_PRESENTATION_J45,
    ('PRESENTATION', 'J47'): calc_PRESENTATION_J47,
    ('PRESENTATION', 'L45'): calc_PRESENTATION_L45,
    ('PRESENTATION', 'L47'): calc_PRESENTATION_L47,
    ('PRESENTATION', 'N45'): calc_PRESENTATION_N45,
    ('PRESENTATION', 'N47'): calc_PRESENTATION_N47,
    ('PRESENTATION', 'P45'): calc_PRESENTATION_P45,
    ('PRESENTATION', 'P47'): calc_PRESENTATION_P47,
    ('PRESENTATION', 'Q45'): calc_PRESENTATION_Q45,
    ('PRESENTATION', 'Q47'): calc_PRESENTATION_Q47,
    ('PRESENTATION', 'S45'): calc_PRESENTATION_S45,
    ('PRESENTATION', 'S47'): calc_PRESENTATION_S47,
    ('PRESENTATION', 'U45'): calc_PRESENTATION_U45,
    ('PRESENTATION', 'U47'): calc_PRESENTATION_U47,
    ('PRESENTATION', 'W45'): calc_PRESENTATION_W45,
    ('PRESENTATION', 'W47'): calc_PRESENTATION_W47,
    ('PRESENTATION', 'X45'): calc_PRESENTATION_X45,
    ('PRESENTATION', 'X47'): calc_PRESENTATION_X47,
    ('PRESENTATION', 'Z45'): calc_PRESENTATION_Z45,
    ('PRESENTATION', 'Z47'): calc_PRESENTATION_Z47,
    ('PRESENTATION', 'AB45'): calc_PRESENTATION_AB45,
    ('PRESENTATION', 'AB47'): calc_PRESENTATION_AB47,
    ('PRESENTATION', 'AD44'): calc_PRESENTATION_AD44,
    ('PRESENTATION', 'AE44'): calc_PRESENTATION_AE44,
    ('PRESENTATION', 'AF44'): calc_PRESENTATION_AF44,
    ('PRESENTATION', 'AG44'): calc_PRESENTATION_AG44,
    ('PRESENTATION', 'H38'): calc_PRESENTATION_H38,
    ('PRESENTATION', 'O38'): calc_PRESENTATION_O38,
    ('PRESENTATION', 'D44'): calc_PRESENTATION_D44,
    ('PRESENTATION', 'F40'): calc_PRESENTATION_F40,
    ('PRESENTATION', 'M40'): calc_PRESENTATION_M40,
    ('PRESENTATION', 'R44'): calc_PRESENTATION_R44,
    ('PRESENTATION', 'T40'): calc_PRESENTATION_T40,
    ('PRESENTATION', 'V38'): calc_PRESENTATION_V38,
    ('PRESENTATION', 'Y44'): calc_PRESENTATION_Y44,
    ('PRESENTATION', 'AA40'): calc_PRESENTATION_AA40,
    ('PRESENTATION', 'AC38'): calc_PRESENTATION_AC38,
    ('Segment Revenue Model', 'U33'): calc_Segment_Revenue_Model_U33,
    ('Segment Revenue Model', 'V33'): calc_Segment_Revenue_Model_V33,
    ('Segment Revenue Model', 'U37'): calc_Segment_Revenue_Model_U37,
    ('Segment Revenue Model', 'U38'): calc_Segment_Revenue_Model_U38,
    ('Segment Revenue Model', 'U39'): calc_Segment_Revenue_Model_U39,
    ('Segment Revenue Model', 'U40'): calc_Segment_Revenue_Model_U40,
    ('Segment Revenue Model', 'U42'): calc_Segment_Revenue_Model_U42,
    ('Segment Revenue Model', 'U43'): calc_Segment_Revenue_Model_U43,
    ('Segment Revenue Model', 'U44'): calc_Segment_Revenue_Model_U44,
    ('Segment Revenue Model', 'U45'): calc_Segment_Revenue_Model_U45,
    ('Segment Revenue Model', 'U46'): calc_Segment_Revenue_Model_U46,
    ('Segment Revenue Model', 'Y15'): calc_Segment_Revenue_Model_Y15,
    ('Segment Revenue Model', 'X17'): calc_Segment_Revenue_Model_X17,
    ('Segment Revenue Model', 'W33'): calc_Segment_Revenue_Model_W33,
    ('Segment Revenue Model', 'W37'): calc_Segment_Revenue_Model_W37,
    ('Segment Revenue Model', 'W38'): calc_Segment_Revenue_Model_W38,
    ('Segment Revenue Model', 'W39'): calc_Segment_Revenue_Model_W39,
    ('Segment Revenue Model', 'W40'): calc_Segment_Revenue_Model_W40,
    ('Segment Revenue Model', 'W41'): calc_Segment_Revenue_Model_W41,
    ('Segment Revenue Model', 'W42'): calc_Segment_Revenue_Model_W42,
    ('Segment Revenue Model', 'W43'): calc_Segment_Revenue_Model_W43,
    ('Segment Revenue Model', 'W44'): calc_Segment_Revenue_Model_W44,
    ('Segment Revenue Model', 'W45'): calc_Segment_Revenue_Model_W45,
    ('Segment Revenue Model', 'W46'): calc_Segment_Revenue_Model_W46,
    ('INCOME STATEMENT', 'B41'): calc_INCOME_STATEMENT_B41,
    ('INCOME STATEMENT', 'B43'): calc_INCOME_STATEMENT_B43,
    ('INCOME STATEMENT', 'B44'): calc_INCOME_STATEMENT_B44,
    ('INCOME STATEMENT', 'C41'): calc_INCOME_STATEMENT_C41,
    ('INCOME STATEMENT', 'C43'): calc_INCOME_STATEMENT_C43,
    ('INCOME STATEMENT', 'C44'): calc_INCOME_STATEMENT_C44,
    ('INCOME STATEMENT', 'E41'): calc_INCOME_STATEMENT_E41,
    ('INCOME STATEMENT', 'E43'): calc_INCOME_STATEMENT_E43,
    ('INCOME STATEMENT', 'E44'): calc_INCOME_STATEMENT_E44,
    ('INCOME STATEMENT', 'G41'): calc_INCOME_STATEMENT_G41,
    ('INCOME STATEMENT', 'G43'): calc_INCOME_STATEMENT_G43,
    ('INCOME STATEMENT', 'G44'): calc_INCOME_STATEMENT_G44,
    ('INCOME STATEMENT', 'I41'): calc_INCOME_STATEMENT_I41,
    ('INCOME STATEMENT', 'I43'): calc_INCOME_STATEMENT_I43,
    ('INCOME STATEMENT', 'I44'): calc_INCOME_STATEMENT_I44,
    ('INCOME STATEMENT', 'J41'): calc_INCOME_STATEMENT_J41,
    ('INCOME STATEMENT', 'J43'): calc_INCOME_STATEMENT_J43,
    ('INCOME STATEMENT', 'J44'): calc_INCOME_STATEMENT_J44,
    ('INCOME STATEMENT', 'L41'): calc_INCOME_STATEMENT_L41,
    ('INCOME STATEMENT', 'L43'): calc_INCOME_STATEMENT_L43,
    ('INCOME STATEMENT', 'L44'): calc_INCOME_STATEMENT_L44,
    ('INCOME STATEMENT', 'N41'): calc_INCOME_STATEMENT_N41,
    ('INCOME STATEMENT', 'N43'): calc_INCOME_STATEMENT_N43,
    ('INCOME STATEMENT', 'N44'): calc_INCOME_STATEMENT_N44,
    ('INCOME STATEMENT', 'P41'): calc_INCOME_STATEMENT_P41,
    ('INCOME STATEMENT', 'P43'): calc_INCOME_STATEMENT_P43,
    ('INCOME STATEMENT', 'P44'): calc_INCOME_STATEMENT_P44,
    ('INCOME STATEMENT', 'Q41'): calc_INCOME_STATEMENT_Q41,
    ('INCOME STATEMENT', 'Q43'): calc_INCOME_STATEMENT_Q43,
    ('INCOME STATEMENT', 'Q44'): calc_INCOME_STATEMENT_Q44,
    ('INCOME STATEMENT', 'S41'): calc_INCOME_STATEMENT_S41,
    ('INCOME STATEMENT', 'S43'): calc_INCOME_STATEMENT_S43,
    ('INCOME STATEMENT', 'S44'): calc_INCOME_STATEMENT_S44,
    ('INCOME STATEMENT', 'U41'): calc_INCOME_STATEMENT_U41,
    ('INCOME STATEMENT', 'U43'): calc_INCOME_STATEMENT_U43,
    ('INCOME STATEMENT', 'U44'): calc_INCOME_STATEMENT_U44,
    ('INCOME STATEMENT', 'W41'): calc_INCOME_STATEMENT_W41,
    ('INCOME STATEMENT', 'W43'): calc_INCOME_STATEMENT_W43,
    ('INCOME STATEMENT', 'W44'): calc_INCOME_STATEMENT_W44,
    ('INCOME STATEMENT', 'X41'): calc_INCOME_STATEMENT_X41,
    ('INCOME STATEMENT', 'X43'): calc_INCOME_STATEMENT_X43,
    ('INCOME STATEMENT', 'X44'): calc_INCOME_STATEMENT_X44,
    ('INCOME STATEMENT', 'Z41'): calc_INCOME_STATEMENT_Z41,
    ('INCOME STATEMENT', 'Z43'): calc_INCOME_STATEMENT_Z43,
    ('INCOME STATEMENT', 'Z44'): calc_INCOME_STATEMENT_Z44,
    ('INCOME STATEMENT', 'AB41'): calc_INCOME_STATEMENT_AB41,
    ('INCOME STATEMENT', 'AB43'): calc_INCOME_STATEMENT_AB43,
    ('INCOME STATEMENT', 'AB44'): calc_INCOME_STATEMENT_AB44,
    ('INCOME STATEMENT', 'H22'): calc_INCOME_STATEMENT_H22,
    ('INCOME STATEMENT', 'H30'): calc_INCOME_STATEMENT_H30,
    ('PRESENTATION', 'C63'): calc_PRESENTATION_C63,
    ('COMPANY OVERVIEW', 'C24'): calc_COMPANY_OVERVIEW_C24,
    ('PRESENTATION', 'C66'): calc_PRESENTATION_C66,
    ('INCOME STATEMENT', 'O22'): calc_INCOME_STATEMENT_O22,
    ('INCOME STATEMENT', 'O30'): calc_INCOME_STATEMENT_O30,
    ('PRESENTATION', 'D63'): calc_PRESENTATION_D63,
    ('COMPANY OVERVIEW', 'D24'): calc_COMPANY_OVERVIEW_D24,
    ('PRESENTATION', 'D66'): calc_PRESENTATION_D66,
    ('INCOME STATEMENT', 'D36'): calc_INCOME_STATEMENT_D36,
    ('INCOME STATEMENT', 'D37'): calc_INCOME_STATEMENT_D37,
    ('INCOME STATEMENT', 'D38'): calc_INCOME_STATEMENT_D38,
    ('INCOME STATEMENT', 'F32'): calc_INCOME_STATEMENT_F32,
    ('INCOME STATEMENT', 'M32'): calc_INCOME_STATEMENT_M32,
    ('INCOME STATEMENT', 'R36'): calc_INCOME_STATEMENT_R36,
    ('INCOME STATEMENT', 'R37'): calc_INCOME_STATEMENT_R37,
    ('INCOME STATEMENT', 'R38'): calc_INCOME_STATEMENT_R38,
    ('INCOME STATEMENT', 'T32'): calc_INCOME_STATEMENT_T32,
    ('INCOME STATEMENT', 'V22'): calc_INCOME_STATEMENT_V22,
    ('INCOME STATEMENT', 'V30'): calc_INCOME_STATEMENT_V30,
    ('Valuation', 'B43'): calc_Valuation_B43,
    ('Valuation', 'B23'): calc_Valuation_B23,
    ('Valuation', 'B44'): calc_Valuation_B44,
    ('Valuation', 'B48'): calc_Valuation_B48,
    ('PRESENTATION', 'E63'): calc_PRESENTATION_E63,
    ('COMPANY OVERVIEW', 'E24'): calc_COMPANY_OVERVIEW_E24,
    ('PRESENTATION', 'E66'): calc_PRESENTATION_E66,
    ('INCOME STATEMENT', 'Y36'): calc_INCOME_STATEMENT_Y36,
    ('INCOME STATEMENT', 'Y37'): calc_INCOME_STATEMENT_Y37,
    ('INCOME STATEMENT', 'Y38'): calc_INCOME_STATEMENT_Y38,
    ('INCOME STATEMENT', 'AA32'): calc_INCOME_STATEMENT_AA32,
    ('INCOME STATEMENT', 'AC22'): calc_INCOME_STATEMENT_AC22,
    ('INCOME STATEMENT', 'AC30'): calc_INCOME_STATEMENT_AC30,
    ('Valuation', 'C21'): calc_Valuation_C21,
    ('Valuation', 'C38'): calc_Valuation_C38,
    ('Valuation', 'D38'): calc_Valuation_D38,
    ('Valuation', 'C43'): calc_Valuation_C43,
    ('PRESENTATION', 'F63'): calc_PRESENTATION_F63,
    ('COMPANY OVERVIEW', 'F24'): calc_COMPANY_OVERVIEW_F24,
    ('PRESENTATION', 'F66'): calc_PRESENTATION_F66,
    ('INCOME STATEMENT', 'AD33'): calc_INCOME_STATEMENT_AD33,
    ('Ratio Analysis', 'G12'): calc_Ratio_Analysis_G12,
    ('Ratio Analysis', 'G18'): calc_Ratio_Analysis_G18,
    ('INCOME STATEMENT', 'AE33'): calc_INCOME_STATEMENT_AE33,
    ('CASH FOW STATEMENT', 'F8'): calc_CASH_FOW_STATEMENT_F8,
    ('Ratio Analysis', 'H12'): calc_Ratio_Analysis_H12,
    ('Ratio Analysis', 'H18'): calc_Ratio_Analysis_H18,
    ('Valuation', 'F44'): calc_Valuation_F44,
    ('INCOME STATEMENT', 'AF33'): calc_INCOME_STATEMENT_AF33,
    ('CASH FOW STATEMENT', 'G8'): calc_CASH_FOW_STATEMENT_G8,
    ('Ratio Analysis', 'I12'): calc_Ratio_Analysis_I12,
    ('Ratio Analysis', 'I18'): calc_Ratio_Analysis_I18,
    ('INCOME STATEMENT', 'AG33'): calc_INCOME_STATEMENT_AG33,
    ('CASH FOW STATEMENT', 'H8'): calc_CASH_FOW_STATEMENT_H8,
    ('Ratio Analysis', 'J12'): calc_Ratio_Analysis_J12,
    ('Ratio Analysis', 'J18'): calc_Ratio_Analysis_J18,
    ('INCOME STATEMENT', 'K45'): calc_INCOME_STATEMENT_K45,
    ('INCOME STATEMENT', 'I47'): calc_INCOME_STATEMENT_I47,
    ('INCOME STATEMENT', 'H48'): calc_INCOME_STATEMENT_H48,
    ('INCOME STATEMENT', 'Z47'): calc_INCOME_STATEMENT_Z47,
    ('INCOME STATEMENT', 'Y48'): calc_INCOME_STATEMENT_Y48,
    ('Valuation', 'I46'): calc_Valuation_I46,
    ('Valuation', 'H24'): calc_Valuation_H24,
    ('Valuation', 'H27'): calc_Valuation_H27,
    ('Valuation', 'I19'): calc_Valuation_I19,
    ('Valuation', 'J45'): calc_Valuation_J45,
    ('PRESENTATION', 'B47'): calc_PRESENTATION_B47,
    ('PRESENTATION', 'AD45'): calc_PRESENTATION_AD45,
    ('PRESENTATION', 'AD47'): calc_PRESENTATION_AD47,
    ('PRESENTATION', 'AE45'): calc_PRESENTATION_AE45,
    ('PRESENTATION', 'AE47'): calc_PRESENTATION_AE47,
    ('PRESENTATION', 'AF45'): calc_PRESENTATION_AF45,
    ('PRESENTATION', 'AF47'): calc_PRESENTATION_AF47,
    ('PRESENTATION', 'AG45'): calc_PRESENTATION_AG45,
    ('PRESENTATION', 'AG47'): calc_PRESENTATION_AG47,
    ('PRESENTATION', 'H40'): calc_PRESENTATION_H40,
    ('PRESENTATION', 'O40'): calc_PRESENTATION_O40,
    ('PRESENTATION', 'D45'): calc_PRESENTATION_D45,
    ('PRESENTATION', 'D47'): calc_PRESENTATION_D47,
    ('PRESENTATION', 'F44'): calc_PRESENTATION_F44,
    ('PRESENTATION', 'M44'): calc_PRESENTATION_M44,
    ('PRESENTATION', 'R45'): calc_PRESENTATION_R45,
    ('PRESENTATION', 'R47'): calc_PRESENTATION_R47,
    ('PRESENTATION', 'T44'): calc_PRESENTATION_T44,
    ('PRESENTATION', 'V40'): calc_PRESENTATION_V40,
    ('PRESENTATION', 'Y45'): calc_PRESENTATION_Y45,
    ('PRESENTATION', 'Y47'): calc_PRESENTATION_Y47,
    ('PRESENTATION', 'AA44'): calc_PRESENTATION_AA44,
    ('PRESENTATION', 'AC40'): calc_PRESENTATION_AC40,
    ('Segment Revenue Model', 'Y17'): calc_Segment_Revenue_Model_Y17,
    ('Segment Revenue Model', 'X33'): calc_Segment_Revenue_Model_X33,
    ('Segment Revenue Model', 'X37'): calc_Segment_Revenue_Model_X37,
    ('Segment Revenue Model', 'X38'): calc_Segment_Revenue_Model_X38,
    ('Segment Revenue Model', 'X39'): calc_Segment_Revenue_Model_X39,
    ('Segment Revenue Model', 'X40'): calc_Segment_Revenue_Model_X40,
    ('Segment Revenue Model', 'X41'): calc_Segment_Revenue_Model_X41,
    ('Segment Revenue Model', 'X42'): calc_Segment_Revenue_Model_X42,
    ('Segment Revenue Model', 'X43'): calc_Segment_Revenue_Model_X43,
    ('Segment Revenue Model', 'X44'): calc_Segment_Revenue_Model_X44,
    ('Segment Revenue Model', 'X45'): calc_Segment_Revenue_Model_X45,
    ('Segment Revenue Model', 'X46'): calc_Segment_Revenue_Model_X46,
    ('INCOME STATEMENT', 'B45'): calc_INCOME_STATEMENT_B45,
    ('INCOME STATEMENT', 'B49'): calc_INCOME_STATEMENT_B49,
    ('INCOME STATEMENT', 'C45'): calc_INCOME_STATEMENT_C45,
    ('INCOME STATEMENT', 'C49'): calc_INCOME_STATEMENT_C49,
    ('INCOME STATEMENT', 'E45'): calc_INCOME_STATEMENT_E45,
    ('INCOME STATEMENT', 'E49'): calc_INCOME_STATEMENT_E49,
    ('INCOME STATEMENT', 'G45'): calc_INCOME_STATEMENT_G45,
    ('INCOME STATEMENT', 'G49'): calc_INCOME_STATEMENT_G49,
    ('INCOME STATEMENT', 'I45'): calc_INCOME_STATEMENT_I45,
    ('INCOME STATEMENT', 'J45'): calc_INCOME_STATEMENT_J45,
    ('INCOME STATEMENT', 'L45'): calc_INCOME_STATEMENT_L45,
    ('INCOME STATEMENT', 'N45'): calc_INCOME_STATEMENT_N45,
    ('INCOME STATEMENT', 'P45'): calc_INCOME_STATEMENT_P45,
    ('INCOME STATEMENT', 'Q45'): calc_INCOME_STATEMENT_Q45,
    ('INCOME STATEMENT', 'S45'): calc_INCOME_STATEMENT_S45,
    ('INCOME STATEMENT', 'S49'): calc_INCOME_STATEMENT_S49,
    ('INCOME STATEMENT', 'U45'): calc_INCOME_STATEMENT_U45,
    ('INCOME STATEMENT', 'U49'): calc_INCOME_STATEMENT_U49,
    ('INCOME STATEMENT', 'W45'): calc_INCOME_STATEMENT_W45,
    ('INCOME STATEMENT', 'W49'): calc_INCOME_STATEMENT_W49,
    ('INCOME STATEMENT', 'X45'): calc_INCOME_STATEMENT_X45,
    ('INCOME STATEMENT', 'X49'): calc_INCOME_STATEMENT_X49,
    ('INCOME STATEMENT', 'Z45'): calc_INCOME_STATEMENT_Z45,
    ('INCOME STATEMENT', 'AB45'): calc_INCOME_STATEMENT_AB45,
    ('INCOME STATEMENT', 'H32'): calc_INCOME_STATEMENT_H32,
    ('INCOME STATEMENT', 'O32'): calc_INCOME_STATEMENT_O32,
    ('Ratio Analysis', 'D26'): calc_Ratio_Analysis_D26,
    ('INCOME STATEMENT', 'D41'): calc_INCOME_STATEMENT_D41,
    ('INCOME STATEMENT', 'D43'): calc_INCOME_STATEMENT_D43,
    ('INCOME STATEMENT', 'K43'): calc_INCOME_STATEMENT_K43,
    ('INCOME STATEMENT', 'F36'): calc_INCOME_STATEMENT_F36,
    ('INCOME STATEMENT', 'F37'): calc_INCOME_STATEMENT_F37,
    ('INCOME STATEMENT', 'F38'): calc_INCOME_STATEMENT_F38,
    ('INCOME STATEMENT', 'M36'): calc_INCOME_STATEMENT_M36,
    ('INCOME STATEMENT', 'M37'): calc_INCOME_STATEMENT_M37,
    ('INCOME STATEMENT', 'M38'): calc_INCOME_STATEMENT_M38,
    ('INCOME STATEMENT', 'R41'): calc_INCOME_STATEMENT_R41,
    ('INCOME STATEMENT', 'R43'): calc_INCOME_STATEMENT_R43,
    ('INCOME STATEMENT', 'T36'): calc_INCOME_STATEMENT_T36,
    ('INCOME STATEMENT', 'T37'): calc_INCOME_STATEMENT_T37,
    ('INCOME STATEMENT', 'T38'): calc_INCOME_STATEMENT_T38,
    ('INCOME STATEMENT', 'V32'): calc_INCOME_STATEMENT_V32,
    ('Ratio Analysis', 'E26'): calc_Ratio_Analysis_E26,
    ('Valuation', 'B26'): calc_Valuation_B26,
    ('INCOME STATEMENT', 'Y41'): calc_INCOME_STATEMENT_Y41,
    ('INCOME STATEMENT', 'Y43'): calc_INCOME_STATEMENT_Y43,
    ('INCOME STATEMENT', 'AA36'): calc_INCOME_STATEMENT_AA36,
    ('INCOME STATEMENT', 'AA37'): calc_INCOME_STATEMENT_AA37,
    ('INCOME STATEMENT', 'AA38'): calc_INCOME_STATEMENT_AA38,
    ('INCOME STATEMENT', 'AC32'): calc_INCOME_STATEMENT_AC32,
    ('Ratio Analysis', 'F26'): calc_Ratio_Analysis_F26,
    ('Valuation', 'C23'): calc_Valuation_C23,
    ('Valuation', 'C39'): calc_Valuation_C39,
    ('Valuation', 'D39'): calc_Valuation_D39,
    ('Valuation', 'C44'): calc_Valuation_C44,
    ('Valuation', 'C48'): calc_Valuation_C48,
    ('INCOME STATEMENT', 'AD36'): calc_INCOME_STATEMENT_AD36,
    ('INCOME STATEMENT', 'AD37'): calc_INCOME_STATEMENT_AD37,
    ('INCOME STATEMENT', 'AD38'): calc_INCOME_STATEMENT_AD38,
    ('Valuation', 'D22'): calc_Valuation_D22,
    ('INCOME STATEMENT', 'AE36'): calc_INCOME_STATEMENT_AE36,
    ('INCOME STATEMENT', 'AE37'): calc_INCOME_STATEMENT_AE37,
    ('INCOME STATEMENT', 'AE38'): calc_INCOME_STATEMENT_AE38,
    ('CASH FOW STATEMENT', 'F9'): calc_CASH_FOW_STATEMENT_F9,
    ('Valuation', 'E22'): calc_Valuation_E22,
    ('Valuation', 'G44'): calc_Valuation_G44,
    ('INCOME STATEMENT', 'AF36'): calc_INCOME_STATEMENT_AF36,
    ('INCOME STATEMENT', 'AF37'): calc_INCOME_STATEMENT_AF37,
    ('INCOME STATEMENT', 'AF38'): calc_INCOME_STATEMENT_AF38,
    ('CASH FOW STATEMENT', 'G9'): calc_CASH_FOW_STATEMENT_G9,
    ('Valuation', 'F22'): calc_Valuation_F22,
    ('INCOME STATEMENT', 'AG36'): calc_INCOME_STATEMENT_AG36,
    ('INCOME STATEMENT', 'AG37'): calc_INCOME_STATEMENT_AG37,
    ('INCOME STATEMENT', 'AG38'): calc_INCOME_STATEMENT_AG38,
    ('CASH FOW STATEMENT', 'H9'): calc_CASH_FOW_STATEMENT_H9,
    ('Valuation', 'G22'): calc_Valuation_G22,
    ('INCOME STATEMENT', 'J47'): calc_INCOME_STATEMENT_J47,
    ('INCOME STATEMENT', 'I48'): calc_INCOME_STATEMENT_I48,
    ('INCOME STATEMENT', 'AA47'): calc_INCOME_STATEMENT_AA47,
    ('INCOME STATEMENT', 'Z48'): calc_INCOME_STATEMENT_Z48,
    ('Valuation', 'J46'): calc_Valuation_J46,
    ('Valuation', 'H40'): calc_Valuation_H40,
    ('Valuation', 'J19'): calc_Valuation_J19,
    ('Valuation', 'I24'): calc_Valuation_I24,
    ('Valuation', 'I27'): calc_Valuation_I27,
    ('Valuation', 'K45'): calc_Valuation_K45,
    ('PRESENTATION', 'H44'): calc_PRESENTATION_H44,
    ('PRESENTATION', 'O44'): calc_PRESENTATION_O44,
    ('PRESENTATION', 'F45'): calc_PRESENTATION_F45,
    ('PRESENTATION', 'F47'): calc_PRESENTATION_F47,
    ('PRESENTATION', 'M45'): calc_PRESENTATION_M45,
    ('PRESENTATION', 'M47'): calc_PRESENTATION_M47,
    ('PRESENTATION', 'T45'): calc_PRESENTATION_T45,
    ('PRESENTATION', 'T47'): calc_PRESENTATION_T47,
    ('PRESENTATION', 'V44'): calc_PRESENTATION_V44,
    ('PRESENTATION', 'AA45'): calc_PRESENTATION_AA45,
    ('PRESENTATION', 'AA47'): calc_PRESENTATION_AA47,
    ('PRESENTATION', 'AC44'): calc_PRESENTATION_AC44,
    ('Segment Revenue Model', 'Y33'): calc_Segment_Revenue_Model_Y33,
    ('Segment Revenue Model', 'Y37'): calc_Segment_Revenue_Model_Y37,
    ('Segment Revenue Model', 'Y38'): calc_Segment_Revenue_Model_Y38,
    ('Segment Revenue Model', 'Y39'): calc_Segment_Revenue_Model_Y39,
    ('Segment Revenue Model', 'Y40'): calc_Segment_Revenue_Model_Y40,
    ('Segment Revenue Model', 'Y41'): calc_Segment_Revenue_Model_Y41,
    ('Segment Revenue Model', 'Y42'): calc_Segment_Revenue_Model_Y42,
    ('Segment Revenue Model', 'Y43'): calc_Segment_Revenue_Model_Y43,
    ('Segment Revenue Model', 'Y44'): calc_Segment_Revenue_Model_Y44,
    ('Segment Revenue Model', 'Y45'): calc_Segment_Revenue_Model_Y45,
    ('Segment Revenue Model', 'Y46'): calc_Segment_Revenue_Model_Y46,
    ('INCOME STATEMENT', 'H36'): calc_INCOME_STATEMENT_H36,
    ('INCOME STATEMENT', 'H37'): calc_INCOME_STATEMENT_H37,
    ('INCOME STATEMENT', 'H38'): calc_INCOME_STATEMENT_H38,
    ('CASH FOW STATEMENT', 'B8'): calc_CASH_FOW_STATEMENT_B8,
    ('Ratio Analysis', 'C12'): calc_Ratio_Analysis_C12,
    ('Ratio Analysis', 'C18'): calc_Ratio_Analysis_C18,
    ('INCOME STATEMENT', 'O36'): calc_INCOME_STATEMENT_O36,
    ('INCOME STATEMENT', 'O37'): calc_INCOME_STATEMENT_O37,
    ('INCOME STATEMENT', 'O38'): calc_INCOME_STATEMENT_O38,
    ('CASH FOW STATEMENT', 'C8'): calc_CASH_FOW_STATEMENT_C8,
    ('Ratio Analysis', 'D12'): calc_Ratio_Analysis_D12,
    ('Ratio Analysis', 'D18'): calc_Ratio_Analysis_D18,
    ('INCOME STATEMENT', 'D45'): calc_INCOME_STATEMENT_D45,
    ('INCOME STATEMENT', 'D49'): calc_INCOME_STATEMENT_D49,
    ('INCOME STATEMENT', 'F41'): calc_INCOME_STATEMENT_F41,
    ('INCOME STATEMENT', 'F43'): calc_INCOME_STATEMENT_F43,
    ('INCOME STATEMENT', 'M41'): calc_INCOME_STATEMENT_M41,
    ('INCOME STATEMENT', 'M43'): calc_INCOME_STATEMENT_M43,
    ('INCOME STATEMENT', 'R45'): calc_INCOME_STATEMENT_R45,
    ('INCOME STATEMENT', 'T41'): calc_INCOME_STATEMENT_T41,
    ('INCOME STATEMENT', 'T43'): calc_INCOME_STATEMENT_T43,
    ('INCOME STATEMENT', 'V36'): calc_INCOME_STATEMENT_V36,
    ('INCOME STATEMENT', 'V37'): calc_INCOME_STATEMENT_V37,
    ('INCOME STATEMENT', 'V38'): calc_INCOME_STATEMENT_V38,
    ('CASH FOW STATEMENT', 'D8'): calc_CASH_FOW_STATEMENT_D8,
    ('Ratio Analysis', 'E12'): calc_Ratio_Analysis_E12,
    ('Ratio Analysis', 'E18'): calc_Ratio_Analysis_E18,
    ('Valuation', 'B28'): calc_Valuation_B28,
    ('INCOME STATEMENT', 'Y45'): calc_INCOME_STATEMENT_Y45,
    ('INCOME STATEMENT', 'Y49'): calc_INCOME_STATEMENT_Y49,
    ('INCOME STATEMENT', 'AA41'): calc_INCOME_STATEMENT_AA41,
    ('INCOME STATEMENT', 'AA43'): calc_INCOME_STATEMENT_AA43,
    ('INCOME STATEMENT', 'AC36'): calc_INCOME_STATEMENT_AC36,
    ('INCOME STATEMENT', 'AC37'): calc_INCOME_STATEMENT_AC37,
    ('INCOME STATEMENT', 'AC38'): calc_INCOME_STATEMENT_AC38,
    ('CASH FOW STATEMENT', 'E8'): calc_CASH_FOW_STATEMENT_E8,
    ('Ratio Analysis', 'F12'): calc_Ratio_Analysis_F12,
    ('Ratio Analysis', 'F18'): calc_Ratio_Analysis_F18,
    ('Valuation', 'C26'): calc_Valuation_C26,
    ('Ratio Analysis', 'G25'): calc_Ratio_Analysis_G25,
    ('Ratio Analysis', 'G16'): calc_Ratio_Analysis_G16,
    ('Ratio Analysis', 'G11'): calc_Ratio_Analysis_G11,
    ('Valuation', 'D23'): calc_Valuation_D23,
    ('Valuation', 'D48'): calc_Valuation_D48,
    ('Ratio Analysis', 'H25'): calc_Ratio_Analysis_H25,
    ('Ratio Analysis', 'H16'): calc_Ratio_Analysis_H16,
    ('Ratio Analysis', 'H11'): calc_Ratio_Analysis_H11,
    ('CASH FOW STATEMENT', 'F10'): calc_CASH_FOW_STATEMENT_F10,
    ('Valuation', 'E23'): calc_Valuation_E23,
    ('Valuation', 'E48'): calc_Valuation_E48,
    ('Valuation', 'H44'): calc_Valuation_H44,
    ('Ratio Analysis', 'I25'): calc_Ratio_Analysis_I25,
    ('Ratio Analysis', 'I16'): calc_Ratio_Analysis_I16,
    ('Ratio Analysis', 'I11'): calc_Ratio_Analysis_I11,
    ('CASH FOW STATEMENT', 'G10'): calc_CASH_FOW_STATEMENT_G10,
    ('Valuation', 'F23'): calc_Valuation_F23,
    ('Ratio Analysis', 'J25'): calc_Ratio_Analysis_J25,
    ('Ratio Analysis', 'J16'): calc_Ratio_Analysis_J16,
    ('Ratio Analysis', 'J11'): calc_Ratio_Analysis_J11,
    ('CASH FOW STATEMENT', 'H10'): calc_CASH_FOW_STATEMENT_H10,
    ('Valuation', 'G23'): calc_Valuation_G23,
    ('INCOME STATEMENT', 'K47'): calc_INCOME_STATEMENT_K47,
    ('INCOME STATEMENT', 'J48'): calc_INCOME_STATEMENT_J48,
    ('INCOME STATEMENT', 'I49'): calc_INCOME_STATEMENT_I49,
    ('INCOME STATEMENT', 'AB47'): calc_INCOME_STATEMENT_AB47,
    ('INCOME STATEMENT', 'AA48'): calc_INCOME_STATEMENT_AA48,
    ('INCOME STATEMENT', 'Z49'): calc_INCOME_STATEMENT_Z49,
    ('Valuation', 'K46'): calc_Valuation_K46,
    ('Valuation', 'K19'): calc_Valuation_K19,
    ('Valuation', 'J24'): calc_Valuation_J24,
    ('Valuation', 'J27'): calc_Valuation_J27,
    ('Valuation', 'I40'): calc_Valuation_I40,
    ('PRESENTATION', 'H45'): calc_PRESENTATION_H45,
    ('PRESENTATION', 'H47'): calc_PRESENTATION_H47,
    ('PRESENTATION', 'O45'): calc_PRESENTATION_O45,
    ('PRESENTATION', 'O47'): calc_PRESENTATION_O47,
    ('PRESENTATION', 'V45'): calc_PRESENTATION_V45,
    ('PRESENTATION', 'V47'): calc_PRESENTATION_V47,
    ('PRESENTATION', 'AC45'): calc_PRESENTATION_AC45,
    ('PRESENTATION', 'AC47'): calc_PRESENTATION_AC47,
    ('Ratio Analysis', 'C25'): calc_Ratio_Analysis_C25,
    ('Ratio Analysis', 'C16'): calc_Ratio_Analysis_C16,
    ('INCOME STATEMENT', 'H41'): calc_INCOME_STATEMENT_H41,
    ('INCOME STATEMENT', 'B42'): calc_INCOME_STATEMENT_B42,
    ('INCOME STATEMENT', 'C42'): calc_INCOME_STATEMENT_C42,
    ('INCOME STATEMENT', 'D42'): calc_INCOME_STATEMENT_D42,
    ('INCOME STATEMENT', 'E42'): calc_INCOME_STATEMENT_E42,
    ('INCOME STATEMENT', 'F42'): calc_INCOME_STATEMENT_F42,
    ('INCOME STATEMENT', 'G42'): calc_INCOME_STATEMENT_G42,
    ('INCOME STATEMENT', 'H43'): calc_INCOME_STATEMENT_H43,
    ('Ratio Analysis', 'C11'): calc_Ratio_Analysis_C11,
    ('CASH FOW STATEMENT', 'B10'): calc_CASH_FOW_STATEMENT_B10,
    ('Ratio Analysis', 'D25'): calc_Ratio_Analysis_D25,
    ('Ratio Analysis', 'D16'): calc_Ratio_Analysis_D16,
    ('INCOME STATEMENT', 'O41'): calc_INCOME_STATEMENT_O41,
    ('INCOME STATEMENT', 'O43'): calc_INCOME_STATEMENT_O43,
    ('Ratio Analysis', 'D11'): calc_Ratio_Analysis_D11,
    ('CASH FOW STATEMENT', 'C10'): calc_CASH_FOW_STATEMENT_C10,
    ('INCOME STATEMENT', 'F45'): calc_INCOME_STATEMENT_F45,
    ('INCOME STATEMENT', 'F49'): calc_INCOME_STATEMENT_F49,
    ('INCOME STATEMENT', 'M45'): calc_INCOME_STATEMENT_M45,
    ('INCOME STATEMENT', 'T45'): calc_INCOME_STATEMENT_T45,
    ('INCOME STATEMENT', 'T49'): calc_INCOME_STATEMENT_T49,
    ('Ratio Analysis', 'E25'): calc_Ratio_Analysis_E25,
    ('Valuation', 'K10'): calc_Valuation_K10,
    ('Ratio Analysis', 'E16'): calc_Ratio_Analysis_E16,
    ('INCOME STATEMENT', 'V41'): calc_INCOME_STATEMENT_V41,
    ('INCOME STATEMENT', 'P42'): calc_INCOME_STATEMENT_P42,
    ('INCOME STATEMENT', 'Q42'): calc_INCOME_STATEMENT_Q42,
    ('INCOME STATEMENT', 'R42'): calc_INCOME_STATEMENT_R42,
    ('INCOME STATEMENT', 'S42'): calc_INCOME_STATEMENT_S42,
    ('INCOME STATEMENT', 'T42'): calc_INCOME_STATEMENT_T42,
    ('INCOME STATEMENT', 'U42'): calc_INCOME_STATEMENT_U42,
    ('INCOME STATEMENT', 'V43'): calc_INCOME_STATEMENT_V43,
    ('Ratio Analysis', 'E11'): calc_Ratio_Analysis_E11,
    ('CASH FOW STATEMENT', 'D10'): calc_CASH_FOW_STATEMENT_D10,
    ('INCOME STATEMENT', 'AC41'): calc_INCOME_STATEMENT_AC41,
    ('INCOME STATEMENT', 'AA45'): calc_INCOME_STATEMENT_AA45,
    ('Ratio Analysis', 'F25'): calc_Ratio_Analysis_F25,
    ('Ratio Analysis', 'F16'): calc_Ratio_Analysis_F16,
    ('INCOME STATEMENT', 'AC43'): calc_INCOME_STATEMENT_AC43,
    ('Ratio Analysis', 'F11'): calc_Ratio_Analysis_F11,
    ('CASH FOW STATEMENT', 'E10'): calc_CASH_FOW_STATEMENT_E10,
    ('Valuation', 'C28'): calc_Valuation_C28,
    ('Valuation', 'D26'): calc_Valuation_D26,
    ('CASH FOW STATEMENT', 'F12'): calc_CASH_FOW_STATEMENT_F12,
    ('Valuation', 'E26'): calc_Valuation_E26,
    ('Valuation', 'F48'): calc_Valuation_F48,
    ('Valuation', 'H21'): calc_Valuation_H21,
    ('Valuation', 'I44'): calc_Valuation_I44,
    ('CASH FOW STATEMENT', 'G12'): calc_CASH_FOW_STATEMENT_G12,
    ('Valuation', 'F26'): calc_Valuation_F26,
    ('CASH FOW STATEMENT', 'H12'): calc_CASH_FOW_STATEMENT_H12,
    ('Valuation', 'G26'): calc_Valuation_G26,
    ('INCOME STATEMENT', 'L47'): calc_INCOME_STATEMENT_L47,
    ('INCOME STATEMENT', 'K48'): calc_INCOME_STATEMENT_K48,
    ('INCOME STATEMENT', 'J49'): calc_INCOME_STATEMENT_J49,
    ('INCOME STATEMENT', 'AC47'): calc_INCOME_STATEMENT_AC47,
    ('INCOME STATEMENT', 'AB48'): calc_INCOME_STATEMENT_AB48,
    ('INCOME STATEMENT', 'AA49'): calc_INCOME_STATEMENT_AA49,
    ('Valuation', 'K24'): calc_Valuation_K24,
    ('Valuation', 'K27'): calc_Valuation_K27,
    ('Valuation', 'J40'): calc_Valuation_J40,
    ('INCOME STATEMENT', 'H45'): calc_INCOME_STATEMENT_H45,
    ('INCOME STATEMENT', 'H49'): calc_INCOME_STATEMENT_H49,
    ('Ratio Analysis', 'C14'): calc_Ratio_Analysis_C14,
    ('INCOME STATEMENT', 'H42'): calc_INCOME_STATEMENT_H42,
    ('CASH FOW STATEMENT', 'B12'): calc_CASH_FOW_STATEMENT_B12,
    ('COMPANY OVERVIEW', 'C23'): calc_COMPANY_OVERVIEW_C23,
    ('INCOME STATEMENT', 'I42'): calc_INCOME_STATEMENT_I42,
    ('INCOME STATEMENT', 'J42'): calc_INCOME_STATEMENT_J42,
    ('INCOME STATEMENT', 'K42'): calc_INCOME_STATEMENT_K42,
    ('INCOME STATEMENT', 'L42'): calc_INCOME_STATEMENT_L42,
    ('INCOME STATEMENT', 'M42'): calc_INCOME_STATEMENT_M42,
    ('INCOME STATEMENT', 'N42'): calc_INCOME_STATEMENT_N42,
    ('INCOME STATEMENT', 'O45'): calc_INCOME_STATEMENT_O45,
    ('Ratio Analysis', 'D14'): calc_Ratio_Analysis_D14,
    ('CASH FOW STATEMENT', 'C12'): calc_CASH_FOW_STATEMENT_C12,
    ('Valuation', 'K13'): calc_Valuation_K13,
    ('COMPANY OVERVIEW', 'D23'): calc_COMPANY_OVERVIEW_D23,
    ('COMPANY OVERVIEW', 'E23'): calc_COMPANY_OVERVIEW_E23,
    ('INCOME STATEMENT', 'V45'): calc_INCOME_STATEMENT_V45,
    ('INCOME STATEMENT', 'V49'): calc_INCOME_STATEMENT_V49,
    ('Ratio Analysis', 'E14'): calc_Ratio_Analysis_E14,
    ('INCOME STATEMENT', 'V42'): calc_INCOME_STATEMENT_V42,
    ('CASH FOW STATEMENT', 'D12'): calc_CASH_FOW_STATEMENT_D12,
    ('COMPANY OVERVIEW', 'F23'): calc_COMPANY_OVERVIEW_F23,
    ('INCOME STATEMENT', 'W42'): calc_INCOME_STATEMENT_W42,
    ('INCOME STATEMENT', 'X42'): calc_INCOME_STATEMENT_X42,
    ('INCOME STATEMENT', 'Y42'): calc_INCOME_STATEMENT_Y42,
    ('INCOME STATEMENT', 'Z42'): calc_INCOME_STATEMENT_Z42,
    ('INCOME STATEMENT', 'AA42'): calc_INCOME_STATEMENT_AA42,
    ('INCOME STATEMENT', 'AB42'): calc_INCOME_STATEMENT_AB42,
    ('INCOME STATEMENT', 'AD43'): calc_INCOME_STATEMENT_AD43,
    ('INCOME STATEMENT', 'AC45'): calc_INCOME_STATEMENT_AC45,
    ('Ratio Analysis', 'F14'): calc_Ratio_Analysis_F14,
    ('CASH FOW STATEMENT', 'E12'): calc_CASH_FOW_STATEMENT_E12,
    ('Valuation', 'D28'): calc_Valuation_D28,
    ('CASH FOW STATEMENT', 'F14'): calc_CASH_FOW_STATEMENT_F14,
    ('Valuation', 'E28'): calc_Valuation_E28,
    ('Valuation', 'G48'): calc_Valuation_G48,
    ('Valuation', 'H20'): calc_Valuation_H20,
    ('Valuation', 'H39'): calc_Valuation_H39,
    ('Valuation', 'I21'): calc_Valuation_I21,
    ('Valuation', 'J44'): calc_Valuation_J44,
    ('CASH FOW STATEMENT', 'G14'): calc_CASH_FOW_STATEMENT_G14,
    ('Valuation', 'F28'): calc_Valuation_F28,
    ('CASH FOW STATEMENT', 'H14'): calc_CASH_FOW_STATEMENT_H14,
    ('Valuation', 'G28'): calc_Valuation_G28,
    ('INCOME STATEMENT', 'M47'): calc_INCOME_STATEMENT_M47,
    ('INCOME STATEMENT', 'L48'): calc_INCOME_STATEMENT_L48,
    ('INCOME STATEMENT', 'K49'): calc_INCOME_STATEMENT_K49,
    ('INCOME STATEMENT', 'AD47'): calc_INCOME_STATEMENT_AD47,
    ('INCOME STATEMENT', 'AC48'): calc_INCOME_STATEMENT_AC48,
    ('INCOME STATEMENT', 'AB49'): calc_INCOME_STATEMENT_AB49,
    ('Valuation', 'K40'): calc_Valuation_K40,
    ('COMPANY OVERVIEW', 'C21'): calc_COMPANY_OVERVIEW_C21,
    ('Ratio Analysis', 'C27'): calc_Ratio_Analysis_C27,
    ('PRESENTATION', 'C65'): calc_PRESENTATION_C65,
    ('CASH FOW STATEMENT', 'B14'): calc_CASH_FOW_STATEMENT_B14,
    ('INCOME STATEMENT', 'O42'): calc_INCOME_STATEMENT_O42,
    ('PRESENTATION', 'D65'): calc_PRESENTATION_D65,
    ('CASH FOW STATEMENT', 'C14'): calc_CASH_FOW_STATEMENT_C14,
    ('Valuation', 'C33'): calc_Valuation_C33,
    ('Valuation', 'D33'): calc_Valuation_D33,
    ('Valuation', 'E33'): calc_Valuation_E33,
    ('Valuation', 'F33'): calc_Valuation_F33,
    ('Valuation', 'G33'): calc_Valuation_G33,
    ('Valuation', 'H33'): calc_Valuation_H33,
    ('Valuation', 'I33'): calc_Valuation_I33,
    ('Valuation', 'J33'): calc_Valuation_J33,
    ('Valuation', 'B53'): calc_Valuation_B53,
    ('COMPANY OVERVIEW', 'E21'): calc_COMPANY_OVERVIEW_E21,
    ('COMPANY OVERVIEW', 'E27'): calc_COMPANY_OVERVIEW_E27,
    ('Ratio Analysis', 'E27'): calc_Ratio_Analysis_E27,
    ('PRESENTATION', 'E65'): calc_PRESENTATION_E65,
    ('CASH FOW STATEMENT', 'D14'): calc_CASH_FOW_STATEMENT_D14,
    ('INCOME STATEMENT', 'AC42'): calc_INCOME_STATEMENT_AC42,
    ('PRESENTATION', 'F65'): calc_PRESENTATION_F65,
    ('CASH FOW STATEMENT', 'E14'): calc_CASH_FOW_STATEMENT_E14,
    ('CASH FOW STATEMENT', 'F23'): calc_CASH_FOW_STATEMENT_F23,
    ('Valuation', 'H48'): calc_Valuation_H48,
    ('Valuation', 'H38'): calc_Valuation_H38,
    ('Valuation', 'I20'): calc_Valuation_I20,
    ('Valuation', 'I39'): calc_Valuation_I39,
    ('Valuation', 'J21'): calc_Valuation_J21,
    ('Valuation', 'K44'): calc_Valuation_K44,
    ('CASH FOW STATEMENT', 'G23'): calc_CASH_FOW_STATEMENT_G23,
    ('CASH FOW STATEMENT', 'H23'): calc_CASH_FOW_STATEMENT_H23,
    ('INCOME STATEMENT', 'N47'): calc_INCOME_STATEMENT_N47,
    ('INCOME STATEMENT', 'M48'): calc_INCOME_STATEMENT_M48,
    ('INCOME STATEMENT', 'L49'): calc_INCOME_STATEMENT_L49,
    ('INCOME STATEMENT', 'AE47'): calc_INCOME_STATEMENT_AE47,
    ('INCOME STATEMENT', 'AD48'): calc_INCOME_STATEMENT_AD48,
    ('INCOME STATEMENT', 'AC49'): calc_INCOME_STATEMENT_AC49,
    ('Valuation', 'C30'): calc_Valuation_C30,
    ('CASH FOW STATEMENT', 'B23'): calc_CASH_FOW_STATEMENT_B23,
    ('CASH FOW STATEMENT', 'C23'): calc_CASH_FOW_STATEMENT_C23,
    ('Valuation', 'C34'): calc_Valuation_C34,
    ('Valuation', 'D34'): calc_Valuation_D34,
    ('Valuation', 'E34'): calc_Valuation_E34,
    ('Valuation', 'F34'): calc_Valuation_F34,
    ('Valuation', 'G34'): calc_Valuation_G34,
    ('Valuation', 'A67'): calc_Valuation_A67,
    ('Valuation', 'A77'): calc_Valuation_A77,
    ('CASH FOW STATEMENT', 'D23'): calc_CASH_FOW_STATEMENT_D23,
    ('CASH FOW STATEMENT', 'E23'): calc_CASH_FOW_STATEMENT_E23,
    ('CASH FOW STATEMENT', 'F24'): calc_CASH_FOW_STATEMENT_F24,
    ('Valuation', 'H22'): calc_Valuation_H22,
    ('Valuation', 'I48'): calc_Valuation_I48,
    ('Valuation', 'I38'): calc_Valuation_I38,
    ('Valuation', 'J20'): calc_Valuation_J20,
    ('Valuation', 'J39'): calc_Valuation_J39,
    ('Valuation', 'K21'): calc_Valuation_K21,
    ('CASH FOW STATEMENT', 'G24'): calc_CASH_FOW_STATEMENT_G24,
    ('CASH FOW STATEMENT', 'H24'): calc_CASH_FOW_STATEMENT_H24,
    ('INCOME STATEMENT', 'O47'): calc_INCOME_STATEMENT_O47,
    ('INCOME STATEMENT', 'P47'): calc_INCOME_STATEMENT_P47,
    ('INCOME STATEMENT', 'N48'): calc_INCOME_STATEMENT_N48,
    ('INCOME STATEMENT', 'M49'): calc_INCOME_STATEMENT_M49,
    ('INCOME STATEMENT', 'AF47'): calc_INCOME_STATEMENT_AF47,
    ('INCOME STATEMENT', 'AE48'): calc_INCOME_STATEMENT_AE48,
    ('INCOME STATEMENT', 'AD49'): calc_INCOME_STATEMENT_AD49,
    ('Valuation', 'D30'): calc_Valuation_D30,
    ('COMPANY OVERVIEW', 'F21'): calc_COMPANY_OVERVIEW_F21,
    ('COMPANY OVERVIEW', 'F27'): calc_COMPANY_OVERVIEW_F27,
    ('Ratio Analysis', 'F27'): calc_Ratio_Analysis_F27,
    ('CASH FOW STATEMENT', 'E24'): calc_CASH_FOW_STATEMENT_E24,
    ('Valuation', 'H23'): calc_Valuation_H23,
    ('Valuation', 'I22'): calc_Valuation_I22,
    ('Valuation', 'J48'): calc_Valuation_J48,
    ('Valuation', 'J38'): calc_Valuation_J38,
    ('Valuation', 'K20'): calc_Valuation_K20,
    ('Valuation', 'K39'): calc_Valuation_K39,
    ('Valuation', 'D75'): calc_Valuation_D75,
    ('Valuation', 'E75'): calc_Valuation_E75,
    ('Valuation', 'F75'): calc_Valuation_F75,
    ('Valuation', 'D76'): calc_Valuation_D76,
    ('Valuation', 'E76'): calc_Valuation_E76,
    ('Valuation', 'F76'): calc_Valuation_F76,
    ('Valuation', 'D77'): calc_Valuation_D77,
    ('Valuation', 'E77'): calc_Valuation_E77,
    ('Valuation', 'F77'): calc_Valuation_F77,
    ('Valuation', 'D78'): calc_Valuation_D78,
    ('Valuation', 'E78'): calc_Valuation_E78,
    ('Valuation', 'F78'): calc_Valuation_F78,
    ('Valuation', 'D79'): calc_Valuation_D79,
    ('Valuation', 'E79'): calc_Valuation_E79,
    ('Valuation', 'F79'): calc_Valuation_F79,
    ('INCOME STATEMENT', 'O48'): calc_INCOME_STATEMENT_O48,
    ('INCOME STATEMENT', 'Q47'): calc_INCOME_STATEMENT_Q47,
    ('INCOME STATEMENT', 'P48'): calc_INCOME_STATEMENT_P48,
    ('INCOME STATEMENT', 'N49'): calc_INCOME_STATEMENT_N49,
    ('INCOME STATEMENT', 'AG47'): calc_INCOME_STATEMENT_AG47,
    ('INCOME STATEMENT', 'AF48'): calc_INCOME_STATEMENT_AF48,
    ('INCOME STATEMENT', 'AE49'): calc_INCOME_STATEMENT_AE49,
    ('Valuation', 'E30'): calc_Valuation_E30,
    ('COMPANY OVERVIEW', 'G21'): calc_COMPANY_OVERVIEW_G21,
    ('COMPANY OVERVIEW', 'G27'): calc_COMPANY_OVERVIEW_G27,
    ('Ratio Analysis', 'G27'): calc_Ratio_Analysis_G27,
    ('Valuation', 'H26'): calc_Valuation_H26,
    ('Valuation', 'I23'): calc_Valuation_I23,
    ('Valuation', 'J22'): calc_Valuation_J22,
    ('Valuation', 'K48'): calc_Valuation_K48,
    ('Valuation', 'K38'): calc_Valuation_K38,
    ('INCOME STATEMENT', 'O49'): calc_INCOME_STATEMENT_O49,
    ('INCOME STATEMENT', 'R47'): calc_INCOME_STATEMENT_R47,
    ('INCOME STATEMENT', 'Q48'): calc_INCOME_STATEMENT_Q48,
    ('INCOME STATEMENT', 'P49'): calc_INCOME_STATEMENT_P49,
    ('INCOME STATEMENT', 'AG48'): calc_INCOME_STATEMENT_AG48,
    ('INCOME STATEMENT', 'AF49'): calc_INCOME_STATEMENT_AF49,
    ('Valuation', 'F30'): calc_Valuation_F30,
    ('COMPANY OVERVIEW', 'H21'): calc_COMPANY_OVERVIEW_H21,
    ('COMPANY OVERVIEW', 'H27'): calc_COMPANY_OVERVIEW_H27,
    ('Ratio Analysis', 'H27'): calc_Ratio_Analysis_H27,
    ('Valuation', 'H28'): calc_Valuation_H28,
    ('Valuation', 'I26'): calc_Valuation_I26,
    ('Valuation', 'J23'): calc_Valuation_J23,
    ('Valuation', 'K22'): calc_Valuation_K22,
    ('COMPANY OVERVIEW', 'D21'): calc_COMPANY_OVERVIEW_D21,
    ('Ratio Analysis', 'D27'): calc_Ratio_Analysis_D27,
    ('INCOME STATEMENT', 'R48'): calc_INCOME_STATEMENT_R48,
    ('INCOME STATEMENT', 'Q49'): calc_INCOME_STATEMENT_Q49,
    ('INCOME STATEMENT', 'AG49'): calc_INCOME_STATEMENT_AG49,
    ('Valuation', 'G30'): calc_Valuation_G30,
    ('COMPANY OVERVIEW', 'I21'): calc_COMPANY_OVERVIEW_I21,
    ('COMPANY OVERVIEW', 'I27'): calc_COMPANY_OVERVIEW_I27,
    ('Ratio Analysis', 'I27'): calc_Ratio_Analysis_I27,
    ('Valuation', 'H34'): calc_Valuation_H34,
    ('Valuation', 'I28'): calc_Valuation_I28,
    ('Valuation', 'J26'): calc_Valuation_J26,
    ('Valuation', 'K23'): calc_Valuation_K23,
    ('INCOME STATEMENT', 'R49'): calc_INCOME_STATEMENT_R49,
    ('COMPANY OVERVIEW', 'J21'): calc_COMPANY_OVERVIEW_J21,
    ('COMPANY OVERVIEW', 'J27'): calc_COMPANY_OVERVIEW_J27,
    ('Ratio Analysis', 'J27'): calc_Ratio_Analysis_J27,
    ('Valuation', 'H30'): calc_Valuation_H30,
    ('Valuation', 'I34'): calc_Valuation_I34,
    ('Valuation', 'J28'): calc_Valuation_J28,
    ('Valuation', 'K26'): calc_Valuation_K26,
    ('Valuation', 'I30'): calc_Valuation_I30,
    ('Valuation', 'J34'): calc_Valuation_J34,
    ('Valuation', 'B65'): calc_Valuation_B65,
    ('Valuation', 'B66'): calc_Valuation_B66,
    ('Valuation', 'B67'): calc_Valuation_B67,
    ('Valuation', 'B68'): calc_Valuation_B68,
    ('Valuation', 'B69'): calc_Valuation_B69,
    ('Valuation', 'B75'): calc_Valuation_B75,
    ('Valuation', 'B76'): calc_Valuation_B76,
    ('Valuation', 'B77'): calc_Valuation_B77,
    ('Valuation', 'B78'): calc_Valuation_B78,
    ('Valuation', 'B79'): calc_Valuation_B79,
    ('Valuation', 'K28'): calc_Valuation_K28,
    ('Valuation', 'J30'): calc_Valuation_J30,
    ('Valuation', 'B52'): calc_Valuation_B52,
    ('Valuation', 'H65'): calc_Valuation_H65,
    ('Valuation', 'I65'): calc_Valuation_I65,
    ('Valuation', 'J65'): calc_Valuation_J65,
    ('Valuation', 'H66'): calc_Valuation_H66,
    ('Valuation', 'I66'): calc_Valuation_I66,
    ('Valuation', 'J66'): calc_Valuation_J66,
    ('Valuation', 'H67'): calc_Valuation_H67,
    ('Valuation', 'I67'): calc_Valuation_I67,
    ('Valuation', 'J67'): calc_Valuation_J67,
    ('Valuation', 'H68'): calc_Valuation_H68,
    ('Valuation', 'I68'): calc_Valuation_I68,
    ('Valuation', 'J68'): calc_Valuation_J68,
    ('Valuation', 'H69'): calc_Valuation_H69,
    ('Valuation', 'I69'): calc_Valuation_I69,
    ('Valuation', 'J69'): calc_Valuation_J69,
    ('Valuation', 'H75'): calc_Valuation_H75,
    ('Valuation', 'I75'): calc_Valuation_I75,
    ('Valuation', 'J75'): calc_Valuation_J75,
    ('Valuation', 'H76'): calc_Valuation_H76,
    ('Valuation', 'I76'): calc_Valuation_I76,
    ('Valuation', 'J76'): calc_Valuation_J76,
    ('Valuation', 'H77'): calc_Valuation_H77,
    ('Valuation', 'I77'): calc_Valuation_I77,
    ('Valuation', 'J77'): calc_Valuation_J77,
    ('Valuation', 'H78'): calc_Valuation_H78,
    ('Valuation', 'I78'): calc_Valuation_I78,
    ('Valuation', 'J78'): calc_Valuation_J78,
    ('Valuation', 'H79'): calc_Valuation_H79,
    ('Valuation', 'I79'): calc_Valuation_I79,
    ('Valuation', 'J79'): calc_Valuation_J79,
    ('Valuation', 'B55'): calc_Valuation_B55,
    ('Valuation', 'K30'): calc_Valuation_K30,
    ('Valuation', 'G52'): calc_Valuation_G52,
    ('Valuation', 'B56'): calc_Valuation_B56,
    ('Valuation', 'C64'): calc_Valuation_C64,
    ('Valuation', 'G56'): calc_Valuation_G56,
    ('Valuation', 'K52'): calc_Valuation_K52,
    ('Valuation', 'K55'): calc_Valuation_K55,
}

FORMULA_ORDER = [
    ('COMPANY OVERVIEW', 'G23'),
    ('COMPANY OVERVIEW', 'H23'),
    ('COMPANY OVERVIEW', 'I23'),
    ('COMPANY OVERVIEW', 'J23'),
    ('COMPANY OVERVIEW', 'E28'),
    ('COMPANY OVERVIEW', 'G28'),
    ('COMPANY OVERVIEW', 'H28'),
    ('COMPANY OVERVIEW', 'I28'),
    ('COMPANY OVERVIEW', 'J28'),
    ('COMPANY OVERVIEW', 'B39'),
    ('COMPANY OVERVIEW', 'B57'),
    ('COMPANY OVERVIEW', 'C57'),
    ('COMPANY OVERVIEW', 'D57'),
    ('COMPANY OVERVIEW', 'E57'),
    ('COMPANY OVERVIEW', 'F57'),
    ('COMPANY OVERVIEW', 'G57'),
    ('PRESENTATION', 'F8'),
    ('PRESENTATION', 'K8'),
    ('PRESENTATION', 'P8'),
    ('PRESENTATION', 'T8'),
    ('PRESENTATION', 'F9'),
    ('PRESENTATION', 'K9'),
    ('PRESENTATION', 'P9'),
    ('PRESENTATION', 'T9'),
    ('PRESENTATION', 'F10'),
    ('PRESENTATION', 'K10'),
    ('PRESENTATION', 'P10'),
    ('PRESENTATION', 'T10'),
    ('PRESENTATION', 'F11'),
    ('PRESENTATION', 'K11'),
    ('PRESENTATION', 'P11'),
    ('PRESENTATION', 'T11'),
    ('PRESENTATION', 'F12'),
    ('PRESENTATION', 'K12'),
    ('PRESENTATION', 'P12'),
    ('PRESENTATION', 'T12'),
    ('PRESENTATION', 'F13'),
    ('PRESENTATION', 'K13'),
    ('PRESENTATION', 'P13'),
    ('PRESENTATION', 'T13'),
    ('PRESENTATION', 'F14'),
    ('PRESENTATION', 'K14'),
    ('PRESENTATION', 'P14'),
    ('PRESENTATION', 'T14'),
    ('PRESENTATION', 'B15'),
    ('PRESENTATION', 'C15'),
    ('PRESENTATION', 'D15'),
    ('PRESENTATION', 'E15'),
    ('PRESENTATION', 'G15'),
    ('PRESENTATION', 'H15'),
    ('PRESENTATION', 'I15'),
    ('PRESENTATION', 'J15'),
    ('PRESENTATION', 'L15'),
    ('PRESENTATION', 'M15'),
    ('PRESENTATION', 'N15'),
    ('PRESENTATION', 'O15'),
    ('PRESENTATION', 'Q15'),
    ('PRESENTATION', 'R15'),
    ('PRESENTATION', 'S15'),
    ('PRESENTATION', 'V15'),
    ('PRESENTATION', 'W15'),
    ('PRESENTATION', 'X15'),
    ('PRESENTATION', 'Y15'),
    ('PRESENTATION', 'F16'),
    ('PRESENTATION', 'P16'),
    ('PRESENTATION', 'U16'),
    ('PRESENTATION', 'D25'),
    ('PRESENTATION', 'K25'),
    ('PRESENTATION', 'R25'),
    ('PRESENTATION', 'Y25'),
    ('PRESENTATION', 'B26'),
    ('PRESENTATION', 'C26'),
    ('PRESENTATION', 'E26'),
    ('PRESENTATION', 'G26'),
    ('PRESENTATION', 'I26'),
    ('PRESENTATION', 'J26'),
    ('PRESENTATION', 'L26'),
    ('PRESENTATION', 'N26'),
    ('PRESENTATION', 'P26'),
    ('PRESENTATION', 'Q26'),
    ('PRESENTATION', 'S26'),
    ('PRESENTATION', 'U26'),
    ('PRESENTATION', 'W26'),
    ('PRESENTATION', 'X26'),
    ('PRESENTATION', 'Z26'),
    ('PRESENTATION', 'AB26'),
    ('PRESENTATION', 'AD26'),
    ('PRESENTATION', 'AE26'),
    ('PRESENTATION', 'AF26'),
    ('PRESENTATION', 'AG26'),
    ('PRESENTATION', 'D27'),
    ('PRESENTATION', 'K27'),
    ('PRESENTATION', 'R27'),
    ('PRESENTATION', 'Y27'),
    ('PRESENTATION', 'D28'),
    ('PRESENTATION', 'K28'),
    ('PRESENTATION', 'R28'),
    ('PRESENTATION', 'Y28'),
    ('PRESENTATION', 'D29'),
    ('PRESENTATION', 'K29'),
    ('PRESENTATION', 'R29'),
    ('PRESENTATION', 'Y29'),
    ('PRESENTATION', 'D30'),
    ('PRESENTATION', 'K30'),
    ('PRESENTATION', 'R30'),
    ('PRESENTATION', 'Y30'),
    ('PRESENTATION', 'D31'),
    ('PRESENTATION', 'K31'),
    ('PRESENTATION', 'R31'),
    ('PRESENTATION', 'Y31'),
    ('PRESENTATION', 'K32'),
    ('PRESENTATION', 'D34'),
    ('PRESENTATION', 'K34'),
    ('PRESENTATION', 'R34'),
    ('PRESENTATION', 'Y34'),
    ('PRESENTATION', 'D35'),
    ('PRESENTATION', 'K35'),
    ('PRESENTATION', 'R35'),
    ('PRESENTATION', 'Y35'),
    ('PRESENTATION', 'D37'),
    ('PRESENTATION', 'K37'),
    ('PRESENTATION', 'R37'),
    ('PRESENTATION', 'V37'),
    ('PRESENTATION', 'Y37'),
    ('PRESENTATION', 'M39'),
    ('PRESENTATION', 'R39'),
    ('PRESENTATION', 'Y39'),
    ('PRESENTATION', 'D41'),
    ('PRESENTATION', 'K41'),
    ('PRESENTATION', 'R41'),
    ('PRESENTATION', 'Y41'),
    ('PRESENTATION', 'D42'),
    ('PRESENTATION', 'K42'),
    ('PRESENTATION', 'R42'),
    ('PRESENTATION', 'Y42'),
    ('PRESENTATION', 'D43'),
    ('PRESENTATION', 'K43'),
    ('PRESENTATION', 'R43'),
    ('PRESENTATION', 'Y43'),
    ('Segment Revenue Model', 'F8'),
    ('Segment Revenue Model', 'K8'),
    ('Segment Revenue Model', 'P8'),
    ('Segment Revenue Model', 'F9'),
    ('Segment Revenue Model', 'K9'),
    ('Segment Revenue Model', 'P9'),
    ('Segment Revenue Model', 'F10'),
    ('Segment Revenue Model', 'K10'),
    ('Segment Revenue Model', 'P10'),
    ('Segment Revenue Model', 'F11'),
    ('Segment Revenue Model', 'K11'),
    ('Segment Revenue Model', 'P11'),
    ('Segment Revenue Model', 'F12'),
    ('Segment Revenue Model', 'K12'),
    ('Segment Revenue Model', 'P12'),
    ('Segment Revenue Model', 'U12'),
    ('Segment Revenue Model', 'F13'),
    ('Segment Revenue Model', 'K13'),
    ('Segment Revenue Model', 'P13'),
    ('Segment Revenue Model', 'F14'),
    ('Segment Revenue Model', 'K14'),
    ('Segment Revenue Model', 'P14'),
    ('Segment Revenue Model', 'B15'),
    ('Segment Revenue Model', 'C15'),
    ('Segment Revenue Model', 'D15'),
    ('Segment Revenue Model', 'E15'),
    ('Segment Revenue Model', 'G15'),
    ('Segment Revenue Model', 'H15'),
    ('Segment Revenue Model', 'I15'),
    ('Segment Revenue Model', 'J15'),
    ('Segment Revenue Model', 'L15'),
    ('Segment Revenue Model', 'M15'),
    ('Segment Revenue Model', 'N15'),
    ('Segment Revenue Model', 'O15'),
    ('Segment Revenue Model', 'Q15'),
    ('Segment Revenue Model', 'R15'),
    ('Segment Revenue Model', 'S15'),
    ('Segment Revenue Model', 'F16'),
    ('Segment Revenue Model', 'P16'),
    ('Segment Revenue Model', 'U16'),
    ('Segment Revenue Model', 'G24'),
    ('Segment Revenue Model', 'H24'),
    ('Segment Revenue Model', 'I24'),
    ('Segment Revenue Model', 'J24'),
    ('Segment Revenue Model', 'L24'),
    ('Segment Revenue Model', 'M24'),
    ('Segment Revenue Model', 'N24'),
    ('Segment Revenue Model', 'O24'),
    ('Segment Revenue Model', 'Q24'),
    ('Segment Revenue Model', 'R24'),
    ('Segment Revenue Model', 'S24'),
    ('Segment Revenue Model', 'G25'),
    ('Segment Revenue Model', 'H25'),
    ('Segment Revenue Model', 'I25'),
    ('Segment Revenue Model', 'J25'),
    ('Segment Revenue Model', 'L25'),
    ('Segment Revenue Model', 'M25'),
    ('Segment Revenue Model', 'N25'),
    ('Segment Revenue Model', 'O25'),
    ('Segment Revenue Model', 'Q25'),
    ('Segment Revenue Model', 'R25'),
    ('Segment Revenue Model', 'S25'),
    ('Segment Revenue Model', 'G26'),
    ('Segment Revenue Model', 'H26'),
    ('Segment Revenue Model', 'I26'),
    ('Segment Revenue Model', 'J26'),
    ('Segment Revenue Model', 'L26'),
    ('Segment Revenue Model', 'M26'),
    ('Segment Revenue Model', 'N26'),
    ('Segment Revenue Model', 'O26'),
    ('Segment Revenue Model', 'Q26'),
    ('Segment Revenue Model', 'R26'),
    ('Segment Revenue Model', 'S26'),
    ('Segment Revenue Model', 'G27'),
    ('Segment Revenue Model', 'H27'),
    ('Segment Revenue Model', 'I27'),
    ('Segment Revenue Model', 'J27'),
    ('Segment Revenue Model', 'L27'),
    ('Segment Revenue Model', 'M27'),
    ('Segment Revenue Model', 'N27'),
    ('Segment Revenue Model', 'O27'),
    ('Segment Revenue Model', 'Q27'),
    ('Segment Revenue Model', 'R27'),
    ('Segment Revenue Model', 'S27'),
    ('Segment Revenue Model', 'G28'),
    ('Segment Revenue Model', 'H28'),
    ('Segment Revenue Model', 'I28'),
    ('Segment Revenue Model', 'J28'),
    ('Segment Revenue Model', 'L28'),
    ('Segment Revenue Model', 'M28'),
    ('Segment Revenue Model', 'N28'),
    ('Segment Revenue Model', 'O28'),
    ('Segment Revenue Model', 'Q28'),
    ('Segment Revenue Model', 'R28'),
    ('Segment Revenue Model', 'S28'),
    ('Segment Revenue Model', 'G29'),
    ('Segment Revenue Model', 'H29'),
    ('Segment Revenue Model', 'I29'),
    ('Segment Revenue Model', 'J29'),
    ('Segment Revenue Model', 'L29'),
    ('Segment Revenue Model', 'M29'),
    ('Segment Revenue Model', 'N29'),
    ('Segment Revenue Model', 'O29'),
    ('Segment Revenue Model', 'Q29'),
    ('Segment Revenue Model', 'R29'),
    ('Segment Revenue Model', 'S29'),
    ('Segment Revenue Model', 'G30'),
    ('Segment Revenue Model', 'H30'),
    ('Segment Revenue Model', 'I30'),
    ('Segment Revenue Model', 'J30'),
    ('Segment Revenue Model', 'L30'),
    ('Segment Revenue Model', 'M30'),
    ('Segment Revenue Model', 'N30'),
    ('Segment Revenue Model', 'O30'),
    ('Segment Revenue Model', 'Q30'),
    ('Segment Revenue Model', 'R30'),
    ('Segment Revenue Model', 'S30'),
    ('Segment Revenue Model', 'G32'),
    ('Segment Revenue Model', 'H32'),
    ('Segment Revenue Model', 'I32'),
    ('Segment Revenue Model', 'J32'),
    ('Segment Revenue Model', 'L32'),
    ('Segment Revenue Model', 'M32'),
    ('Segment Revenue Model', 'N32'),
    ('Segment Revenue Model', 'O32'),
    ('Segment Revenue Model', 'Q32'),
    ('Segment Revenue Model', 'R32'),
    ('Segment Revenue Model', 'S32'),
    ('Segment Revenue Model', 'T32'),
    ('Segment Revenue Model', 'F51'),
    ('Segment Revenue Model', 'K51'),
    ('Segment Revenue Model', 'P51'),
    ('Segment Revenue Model', 'U51'),
    ('Segment Revenue Model', 'F52'),
    ('Segment Revenue Model', 'P52'),
    ('Segment Revenue Model', 'U52'),
    ('Segment Revenue Model', 'F53'),
    ('Segment Revenue Model', 'K53'),
    ('Segment Revenue Model', 'P53'),
    ('Segment Revenue Model', 'U53'),
    ('Segment Revenue Model', 'F54'),
    ('Segment Revenue Model', 'K54'),
    ('Segment Revenue Model', 'P54'),
    ('Segment Revenue Model', 'U54'),
    ('Segment Revenue Model', 'F55'),
    ('Segment Revenue Model', 'K55'),
    ('Segment Revenue Model', 'P55'),
    ('Segment Revenue Model', 'U55'),
    ('Segment Revenue Model', 'F56'),
    ('Segment Revenue Model', 'K56'),
    ('Segment Revenue Model', 'P56'),
    ('Segment Revenue Model', 'U56'),
    ('Segment Revenue Model', 'F57'),
    ('Segment Revenue Model', 'K57'),
    ('Segment Revenue Model', 'P57'),
    ('Segment Revenue Model', 'U57'),
    ('Segment Revenue Model', 'B58'),
    ('Segment Revenue Model', 'C58'),
    ('Segment Revenue Model', 'D58'),
    ('Segment Revenue Model', 'E58'),
    ('Segment Revenue Model', 'G58'),
    ('Segment Revenue Model', 'H58'),
    ('Segment Revenue Model', 'I58'),
    ('Segment Revenue Model', 'J58'),
    ('Segment Revenue Model', 'L58'),
    ('Segment Revenue Model', 'M58'),
    ('Segment Revenue Model', 'N58'),
    ('Segment Revenue Model', 'O58'),
    ('Segment Revenue Model', 'Q58'),
    ('Segment Revenue Model', 'R58'),
    ('Segment Revenue Model', 'S58'),
    ('Segment Revenue Model', 'T58'),
    ('Segment Revenue Model', 'V58'),
    ('Segment Revenue Model', 'W58'),
    ('Segment Revenue Model', 'X58'),
    ('Segment Revenue Model', 'Y58'),
    ('Segment Revenue Model', 'F59'),
    ('Segment Revenue Model', 'K59'),
    ('Segment Revenue Model', 'P59'),
    ('Segment Revenue Model', 'U59'),
    ('Segment Revenue Model', 'F60'),
    ('Segment Revenue Model', 'K60'),
    ('Segment Revenue Model', 'P60'),
    ('Segment Revenue Model', 'U60'),
    ('Segment Revenue Model', 'L61'),
    ('Segment Revenue Model', 'F65'),
    ('Segment Revenue Model', 'K65'),
    ('Segment Revenue Model', 'P65'),
    ('Segment Revenue Model', 'U65'),
    ('Segment Revenue Model', 'F66'),
    ('Segment Revenue Model', 'K66'),
    ('Segment Revenue Model', 'P66'),
    ('Segment Revenue Model', 'U66'),
    ('Segment Revenue Model', 'F67'),
    ('Segment Revenue Model', 'K67'),
    ('Segment Revenue Model', 'P67'),
    ('Segment Revenue Model', 'U67'),
    ('Segment Revenue Model', 'F68'),
    ('Segment Revenue Model', 'K68'),
    ('Segment Revenue Model', 'P68'),
    ('Segment Revenue Model', 'U68'),
    ('Segment Revenue Model', 'F69'),
    ('Segment Revenue Model', 'K69'),
    ('Segment Revenue Model', 'P69'),
    ('Segment Revenue Model', 'U69'),
    ('Segment Revenue Model', 'F70'),
    ('Segment Revenue Model', 'K70'),
    ('Segment Revenue Model', 'P70'),
    ('Segment Revenue Model', 'U70'),
    ('Segment Revenue Model', 'F71'),
    ('Segment Revenue Model', 'K71'),
    ('Segment Revenue Model', 'P71'),
    ('Segment Revenue Model', 'U71'),
    ('Segment Revenue Model', 'B72'),
    ('Segment Revenue Model', 'C72'),
    ('Segment Revenue Model', 'D72'),
    ('Segment Revenue Model', 'E72'),
    ('Segment Revenue Model', 'G72'),
    ('Segment Revenue Model', 'H72'),
    ('Segment Revenue Model', 'I72'),
    ('Segment Revenue Model', 'J72'),
    ('Segment Revenue Model', 'L72'),
    ('Segment Revenue Model', 'M72'),
    ('Segment Revenue Model', 'N72'),
    ('Segment Revenue Model', 'O72'),
    ('Segment Revenue Model', 'Q72'),
    ('Segment Revenue Model', 'R72'),
    ('Segment Revenue Model', 'S72'),
    ('Segment Revenue Model', 'T72'),
    ('Segment Revenue Model', 'F73'),
    ('Segment Revenue Model', 'K73'),
    ('Segment Revenue Model', 'P73'),
    ('Segment Revenue Model', 'U73'),
    ('INCOME STATEMENT', 'D6'),
    ('INCOME STATEMENT', 'K6'),
    ('INCOME STATEMENT', 'R6'),
    ('INCOME STATEMENT', 'Y6'),
    ('INCOME STATEMENT', 'B7'),
    ('INCOME STATEMENT', 'C7'),
    ('INCOME STATEMENT', 'E7'),
    ('INCOME STATEMENT', 'G7'),
    ('INCOME STATEMENT', 'P7'),
    ('INCOME STATEMENT', 'Q7'),
    ('INCOME STATEMENT', 'S7'),
    ('INCOME STATEMENT', 'U7'),
    ('INCOME STATEMENT', 'H8'),
    ('INCOME STATEMENT', 'I8'),
    ('INCOME STATEMENT', 'J8'),
    ('INCOME STATEMENT', 'L8'),
    ('INCOME STATEMENT', 'N8'),
    ('INCOME STATEMENT', 'P8'),
    ('INCOME STATEMENT', 'Q8'),
    ('INCOME STATEMENT', 'S8'),
    ('INCOME STATEMENT', 'U8'),
    ('INCOME STATEMENT', 'W8'),
    ('INCOME STATEMENT', 'X8'),
    ('INCOME STATEMENT', 'Z8'),
    ('INCOME STATEMENT', 'AB8'),
    ('INCOME STATEMENT', 'AE8'),
    ('INCOME STATEMENT', 'AF8'),
    ('INCOME STATEMENT', 'AG8'),
    ('INCOME STATEMENT', 'C9'),
    ('INCOME STATEMENT', 'E9'),
    ('INCOME STATEMENT', 'G9'),
    ('INCOME STATEMENT', 'I9'),
    ('INCOME STATEMENT', 'J9'),
    ('INCOME STATEMENT', 'L9'),
    ('INCOME STATEMENT', 'N9'),
    ('INCOME STATEMENT', 'P9'),
    ('INCOME STATEMENT', 'Q9'),
    ('INCOME STATEMENT', 'S9'),
    ('INCOME STATEMENT', 'U9'),
    ('INCOME STATEMENT', 'W9'),
    ('INCOME STATEMENT', 'X9'),
    ('INCOME STATEMENT', 'Z9'),
    ('INCOME STATEMENT', 'AB9'),
    ('INCOME STATEMENT', 'B10'),
    ('INCOME STATEMENT', 'C10'),
    ('INCOME STATEMENT', 'E10'),
    ('INCOME STATEMENT', 'G10'),
    ('INCOME STATEMENT', 'I10'),
    ('INCOME STATEMENT', 'J10'),
    ('INCOME STATEMENT', 'L10'),
    ('INCOME STATEMENT', 'N10'),
    ('INCOME STATEMENT', 'P10'),
    ('INCOME STATEMENT', 'Q10'),
    ('INCOME STATEMENT', 'S10'),
    ('INCOME STATEMENT', 'U10'),
    ('INCOME STATEMENT', 'W10'),
    ('INCOME STATEMENT', 'X10'),
    ('INCOME STATEMENT', 'Z10'),
    ('INCOME STATEMENT', 'AB10'),
    ('INCOME STATEMENT', 'D11'),
    ('INCOME STATEMENT', 'K11'),
    ('INCOME STATEMENT', 'R11'),
    ('INCOME STATEMENT', 'Y11'),
    ('INCOME STATEMENT', 'AD11'),
    ('INCOME STATEMENT', 'AE11'),
    ('INCOME STATEMENT', 'AF11'),
    ('INCOME STATEMENT', 'AG11'),
    ('INCOME STATEMENT', 'B12'),
    ('INCOME STATEMENT', 'C12'),
    ('INCOME STATEMENT', 'E12'),
    ('INCOME STATEMENT', 'G12'),
    ('INCOME STATEMENT', 'H12'),
    ('INCOME STATEMENT', 'I12'),
    ('INCOME STATEMENT', 'J12'),
    ('INCOME STATEMENT', 'L12'),
    ('INCOME STATEMENT', 'N12'),
    ('INCOME STATEMENT', 'P12'),
    ('INCOME STATEMENT', 'Q12'),
    ('INCOME STATEMENT', 'S12'),
    ('INCOME STATEMENT', 'U12'),
    ('INCOME STATEMENT', 'V12'),
    ('INCOME STATEMENT', 'W12'),
    ('INCOME STATEMENT', 'X12'),
    ('INCOME STATEMENT', 'Z12'),
    ('INCOME STATEMENT', 'AB12'),
    ('INCOME STATEMENT', 'D13'),
    ('INCOME STATEMENT', 'K13'),
    ('INCOME STATEMENT', 'R13'),
    ('INCOME STATEMENT', 'Y13'),
    ('INCOME STATEMENT', 'AD13'),
    ('INCOME STATEMENT', 'AE13'),
    ('INCOME STATEMENT', 'AF13'),
    ('INCOME STATEMENT', 'AG13'),
    ('INCOME STATEMENT', 'B14'),
    ('INCOME STATEMENT', 'C14'),
    ('INCOME STATEMENT', 'E14'),
    ('INCOME STATEMENT', 'G14'),
    ('INCOME STATEMENT', 'H14'),
    ('INCOME STATEMENT', 'I14'),
    ('INCOME STATEMENT', 'J14'),
    ('INCOME STATEMENT', 'L14'),
    ('INCOME STATEMENT', 'N14'),
    ('INCOME STATEMENT', 'P14'),
    ('INCOME STATEMENT', 'Q14'),
    ('INCOME STATEMENT', 'S14'),
    ('INCOME STATEMENT', 'U14'),
    ('INCOME STATEMENT', 'V14'),
    ('INCOME STATEMENT', 'W14'),
    ('INCOME STATEMENT', 'X14'),
    ('INCOME STATEMENT', 'Z14'),
    ('INCOME STATEMENT', 'AB14'),
    ('INCOME STATEMENT', 'D15'),
    ('INCOME STATEMENT', 'K15'),
    ('INCOME STATEMENT', 'R15'),
    ('INCOME STATEMENT', 'Y15'),
    ('INCOME STATEMENT', 'AD15'),
    ('INCOME STATEMENT', 'AE15'),
    ('INCOME STATEMENT', 'AF15'),
    ('INCOME STATEMENT', 'AG15'),
    ('INCOME STATEMENT', 'B16'),
    ('INCOME STATEMENT', 'C16'),
    ('INCOME STATEMENT', 'E16'),
    ('INCOME STATEMENT', 'G16'),
    ('INCOME STATEMENT', 'I16'),
    ('INCOME STATEMENT', 'J16'),
    ('INCOME STATEMENT', 'L16'),
    ('INCOME STATEMENT', 'N16'),
    ('INCOME STATEMENT', 'P16'),
    ('INCOME STATEMENT', 'Q16'),
    ('INCOME STATEMENT', 'S16'),
    ('INCOME STATEMENT', 'U16'),
    ('INCOME STATEMENT', 'W16'),
    ('INCOME STATEMENT', 'X16'),
    ('INCOME STATEMENT', 'Z16'),
    ('INCOME STATEMENT', 'AB16'),
    ('INCOME STATEMENT', 'D17'),
    ('INCOME STATEMENT', 'K17'),
    ('INCOME STATEMENT', 'R17'),
    ('INCOME STATEMENT', 'Y17'),
    ('INCOME STATEMENT', 'AD17'),
    ('INCOME STATEMENT', 'AE17'),
    ('INCOME STATEMENT', 'AF17'),
    ('INCOME STATEMENT', 'AG17'),
    ('INCOME STATEMENT', 'B18'),
    ('INCOME STATEMENT', 'C18'),
    ('INCOME STATEMENT', 'E18'),
    ('INCOME STATEMENT', 'G18'),
    ('INCOME STATEMENT', 'H18'),
    ('INCOME STATEMENT', 'I18'),
    ('INCOME STATEMENT', 'J18'),
    ('INCOME STATEMENT', 'L18'),
    ('INCOME STATEMENT', 'N18'),
    ('INCOME STATEMENT', 'P18'),
    ('INCOME STATEMENT', 'Q18'),
    ('INCOME STATEMENT', 'S18'),
    ('INCOME STATEMENT', 'U18'),
    ('INCOME STATEMENT', 'V18'),
    ('INCOME STATEMENT', 'W18'),
    ('INCOME STATEMENT', 'X18'),
    ('INCOME STATEMENT', 'Z18'),
    ('INCOME STATEMENT', 'AB18'),
    ('INCOME STATEMENT', 'D19'),
    ('INCOME STATEMENT', 'K19'),
    ('INCOME STATEMENT', 'R19'),
    ('INCOME STATEMENT', 'Y19'),
    ('INCOME STATEMENT', 'AD19'),
    ('INCOME STATEMENT', 'AE19'),
    ('INCOME STATEMENT', 'AF19'),
    ('INCOME STATEMENT', 'AG19'),
    ('INCOME STATEMENT', 'B20'),
    ('INCOME STATEMENT', 'C20'),
    ('INCOME STATEMENT', 'E20'),
    ('INCOME STATEMENT', 'G20'),
    ('INCOME STATEMENT', 'I20'),
    ('INCOME STATEMENT', 'J20'),
    ('INCOME STATEMENT', 'L20'),
    ('INCOME STATEMENT', 'N20'),
    ('INCOME STATEMENT', 'P20'),
    ('INCOME STATEMENT', 'Q20'),
    ('INCOME STATEMENT', 'S20'),
    ('INCOME STATEMENT', 'U20'),
    ('INCOME STATEMENT', 'W20'),
    ('INCOME STATEMENT', 'X20'),
    ('INCOME STATEMENT', 'Z20'),
    ('INCOME STATEMENT', 'AB20'),
    ('INCOME STATEMENT', 'K21'),
    ('INCOME STATEMENT', 'D26'),
    ('INCOME STATEMENT', 'K26'),
    ('INCOME STATEMENT', 'R26'),
    ('INCOME STATEMENT', 'Y26'),
    ('INCOME STATEMENT', 'D27'),
    ('INCOME STATEMENT', 'K27'),
    ('INCOME STATEMENT', 'R27'),
    ('INCOME STATEMENT', 'Y27'),
    ('INCOME STATEMENT', 'D29'),
    ('INCOME STATEMENT', 'K29'),
    ('INCOME STATEMENT', 'R29'),
    ('INCOME STATEMENT', 'V29'),
    ('INCOME STATEMENT', 'Y29'),
    ('INCOME STATEMENT', 'M31'),
    ('INCOME STATEMENT', 'R31'),
    ('INCOME STATEMENT', 'Y31'),
    ('INCOME STATEMENT', 'D33'),
    ('INCOME STATEMENT', 'K33'),
    ('INCOME STATEMENT', 'R33'),
    ('INCOME STATEMENT', 'Y33'),
    ('INCOME STATEMENT', 'D34'),
    ('INCOME STATEMENT', 'K34'),
    ('INCOME STATEMENT', 'R34'),
    ('INCOME STATEMENT', 'Y34'),
    ('INCOME STATEMENT', 'D35'),
    ('INCOME STATEMENT', 'K35'),
    ('INCOME STATEMENT', 'R35'),
    ('INCOME STATEMENT', 'Y35'),
    ('INCOME STATEMENT', 'D39'),
    ('INCOME STATEMENT', 'H39'),
    ('INCOME STATEMENT', 'D40'),
    ('INCOME STATEMENT', 'AE43'),
    ('INCOME STATEMENT', 'AF43'),
    ('INCOME STATEMENT', 'AG43'),
    ('INCOME STATEMENT', 'AD45'),
    ('INCOME STATEMENT', 'AE45'),
    ('INCOME STATEMENT', 'AF45'),
    ('INCOME STATEMENT', 'AG45'),
    ('INCOME STATEMENT', 'C47'),
    ('INCOME STATEMENT', 'T47'),
    ('INCOME STATEMENT', 'B48'),
    ('INCOME STATEMENT', 'S48'),
    ('INCOME STATEMENT', 'W66'),
    ('CASH FOW STATEMENT', 'I8'),
    ('CASH FOW STATEMENT', 'C9'),
    ('CASH FOW STATEMENT', 'D9'),
    ('CASH FOW STATEMENT', 'I9'),
    ('CASH FOW STATEMENT', 'E11'),
    ('CASH FOW STATEMENT', 'F11'),
    ('CASH FOW STATEMENT', 'G11'),
    ('CASH FOW STATEMENT', 'H11'),
    ('CASH FOW STATEMENT', 'I11'),
    ('CASH FOW STATEMENT', 'B13'),
    ('CASH FOW STATEMENT', 'C13'),
    ('CASH FOW STATEMENT', 'D13'),
    ('CASH FOW STATEMENT', 'E13'),
    ('CASH FOW STATEMENT', 'F13'),
    ('CASH FOW STATEMENT', 'G13'),
    ('CASH FOW STATEMENT', 'H13'),
    ('CASH FOW STATEMENT', 'I13'),
    ('CASH FOW STATEMENT', 'B17'),
    ('CASH FOW STATEMENT', 'C17'),
    ('CASH FOW STATEMENT', 'D17'),
    ('CASH FOW STATEMENT', 'E17'),
    ('CASH FOW STATEMENT', 'F17'),
    ('CASH FOW STATEMENT', 'G17'),
    ('CASH FOW STATEMENT', 'H17'),
    ('CASH FOW STATEMENT', 'I17'),
    ('CASH FOW STATEMENT', 'B19'),
    ('CASH FOW STATEMENT', 'C19'),
    ('CASH FOW STATEMENT', 'D19'),
    ('CASH FOW STATEMENT', 'E19'),
    ('CASH FOW STATEMENT', 'F19'),
    ('CASH FOW STATEMENT', 'G19'),
    ('CASH FOW STATEMENT', 'H19'),
    ('CASH FOW STATEMENT', 'I19'),
    ('CASH FOW STATEMENT', 'F20'),
    ('CASH FOW STATEMENT', 'G20'),
    ('CASH FOW STATEMENT', 'H20'),
    ('CASH FOW STATEMENT', 'I20'),
    ('CASH FOW STATEMENT', 'E22'),
    ('BALANCESHEET', 'B10'),
    ('BALANCESHEET', 'C10'),
    ('BALANCESHEET', 'D10'),
    ('BALANCESHEET', 'E10'),
    ('BALANCESHEET', 'F10'),
    ('BALANCESHEET', 'G10'),
    ('BALANCESHEET', 'H10'),
    ('BALANCESHEET', 'I10'),
    ('BALANCESHEET', 'J10'),
    ('BALANCESHEET', 'B19'),
    ('BALANCESHEET', 'C19'),
    ('BALANCESHEET', 'D19'),
    ('BALANCESHEET', 'E19'),
    ('BALANCESHEET', 'F19'),
    ('BALANCESHEET', 'G19'),
    ('BALANCESHEET', 'H19'),
    ('BALANCESHEET', 'I19'),
    ('BALANCESHEET', 'J19'),
    ('BALANCESHEET', 'B30'),
    ('BALANCESHEET', 'C30'),
    ('BALANCESHEET', 'D30'),
    ('BALANCESHEET', 'E30'),
    ('BALANCESHEET', 'F30'),
    ('BALANCESHEET', 'G30'),
    ('BALANCESHEET', 'H30'),
    ('BALANCESHEET', 'I30'),
    ('BALANCESHEET', 'J30'),
    ('BALANCESHEET', 'B36'),
    ('BALANCESHEET', 'C36'),
    ('BALANCESHEET', 'D36'),
    ('BALANCESHEET', 'E36'),
    ('BALANCESHEET', 'F36'),
    ('BALANCESHEET', 'G36'),
    ('BALANCESHEET', 'H36'),
    ('BALANCESHEET', 'I36'),
    ('BALANCESHEET', 'J36'),
    ('BALANCESHEET', 'B44'),
    ('BALANCESHEET', 'C44'),
    ('BALANCESHEET', 'D44'),
    ('BALANCESHEET', 'E44'),
    ('BALANCESHEET', 'F44'),
    ('BALANCESHEET', 'G44'),
    ('BALANCESHEET', 'H44'),
    ('BALANCESHEET', 'I44'),
    ('BALANCESHEET', 'J44'),
    ('BALANCESHEET', 'B52'),
    ('BALANCESHEET', 'C52'),
    ('BALANCESHEET', 'D52'),
    ('BALANCESHEET', 'E52'),
    ('Debt Schedule', 'B11'),
    ('Debt Schedule', 'C11'),
    ('Debt Schedule', 'D11'),
    ('Debt Schedule', 'E11'),
    ('Debt Schedule', 'F11'),
    ('Debt Schedule', 'G11'),
    ('Debt Schedule', 'H11'),
    ('Debt Schedule', 'I11'),
    ('Debt Schedule', 'B20'),
    ('Debt Schedule', 'C20'),
    ('Debt Schedule', 'D20'),
    ('Debt Schedule', 'E20'),
    ('Debt Schedule', 'F20'),
    ('Debt Schedule', 'G20'),
    ('Debt Schedule', 'H20'),
    ('Debt Schedule', 'I20'),
    ('Valuation', 'G9'),
    ('Valuation', 'B10'),
    ('Valuation', 'G11'),
    ('Valuation', 'B19'),
    ('Valuation', 'D19'),
    ('Valuation', 'E19'),
    ('Valuation', 'F19'),
    ('Valuation', 'G19'),
    ('Valuation', 'B22'),
    ('Valuation', 'C24'),
    ('Valuation', 'D24'),
    ('Valuation', 'E24'),
    ('Valuation', 'F24'),
    ('Valuation', 'G24'),
    ('Valuation', 'K37'),
    ('Valuation', 'K53'),
    ('Valuation', 'B54'),
    ('Valuation', 'H64'),
    ('Valuation', 'I64'),
    ('Valuation', 'J64'),
    ('Valuation', 'H74'),
    ('Valuation', 'I74'),
    ('Valuation', 'J74'),
    ('Ratio Analysis', 'G14'),
    ('Ratio Analysis', 'H14'),
    ('Ratio Analysis', 'I14'),
    ('Ratio Analysis', 'J14'),
    ('Ratio Analysis', 'C20'),
    ('Ratio Analysis', 'D20'),
    ('Ratio Analysis', 'E20'),
    ('Ratio Analysis', 'F20'),
    ('Ratio Analysis', 'G20'),
    ('Ratio Analysis', 'H20'),
    ('Ratio Analysis', 'I20'),
    ('Ratio Analysis', 'J20'),
    ('Ratio Analysis', 'G21'),
    ('Ratio Analysis', 'H21'),
    ('Ratio Analysis', 'J21'),
    ('Ratio Analysis', 'E23'),
    ('Ratio Analysis', 'G23'),
    ('Ratio Analysis', 'H23'),
    ('Ratio Analysis', 'I23'),
    ('Ratio Analysis', 'J23'),
    ('Ratio Analysis', 'G24'),
    ('Ratio Analysis', 'H24'),
    ('Ratio Analysis', 'I24'),
    ('Ratio Analysis', 'J24'),
    ('PRESENTATION', 'F15'),
    ('PRESENTATION', 'K15'),
    ('PRESENTATION', 'P15'),
    ('PRESENTATION', 'T15'),
    ('PRESENTATION', 'B17'),
    ('PRESENTATION', 'C17'),
    ('PRESENTATION', 'D17'),
    ('PRESENTATION', 'E17'),
    ('PRESENTATION', 'G17'),
    ('PRESENTATION', 'H17'),
    ('PRESENTATION', 'I17'),
    ('PRESENTATION', 'J17'),
    ('PRESENTATION', 'L17'),
    ('PRESENTATION', 'M17'),
    ('PRESENTATION', 'N17'),
    ('PRESENTATION', 'O17'),
    ('PRESENTATION', 'Q17'),
    ('PRESENTATION', 'R17'),
    ('PRESENTATION', 'S17'),
    ('PRESENTATION', 'V17'),
    ('PRESENTATION', 'W17'),
    ('PRESENTATION', 'X17'),
    ('PRESENTATION', 'Y17'),
    ('PRESENTATION', 'F25'),
    ('PRESENTATION', 'M25'),
    ('PRESENTATION', 'T25'),
    ('PRESENTATION', 'AA25'),
    ('PRESENTATION', 'B32'),
    ('PRESENTATION', 'C32'),
    ('PRESENTATION', 'E32'),
    ('PRESENTATION', 'G32'),
    ('PRESENTATION', 'I32'),
    ('PRESENTATION', 'J32'),
    ('PRESENTATION', 'L32'),
    ('PRESENTATION', 'N32'),
    ('PRESENTATION', 'P32'),
    ('PRESENTATION', 'Q32'),
    ('PRESENTATION', 'S32'),
    ('PRESENTATION', 'U32'),
    ('PRESENTATION', 'W32'),
    ('PRESENTATION', 'X32'),
    ('PRESENTATION', 'Z32'),
    ('PRESENTATION', 'AB32'),
    ('PRESENTATION', 'AD32'),
    ('PRESENTATION', 'AE32'),
    ('PRESENTATION', 'AF32'),
    ('PRESENTATION', 'AG32'),
    ('PRESENTATION', 'F27'),
    ('PRESENTATION', 'M27'),
    ('PRESENTATION', 'T27'),
    ('PRESENTATION', 'AA27'),
    ('PRESENTATION', 'F28'),
    ('PRESENTATION', 'M28'),
    ('PRESENTATION', 'T28'),
    ('PRESENTATION', 'AA28'),
    ('PRESENTATION', 'F29'),
    ('PRESENTATION', 'M29'),
    ('PRESENTATION', 'T29'),
    ('PRESENTATION', 'AA29'),
    ('PRESENTATION', 'F30'),
    ('PRESENTATION', 'M30'),
    ('PRESENTATION', 'T30'),
    ('PRESENTATION', 'AA30'),
    ('PRESENTATION', 'D26'),
    ('PRESENTATION', 'F31'),
    ('PRESENTATION', 'M31'),
    ('PRESENTATION', 'R26'),
    ('PRESENTATION', 'T31'),
    ('PRESENTATION', 'Y26'),
    ('PRESENTATION', 'AA31'),
    ('PRESENTATION', 'K33'),
    ('PRESENTATION', 'F34'),
    ('PRESENTATION', 'M34'),
    ('PRESENTATION', 'T34'),
    ('PRESENTATION', 'AA34'),
    ('PRESENTATION', 'F35'),
    ('PRESENTATION', 'M35'),
    ('PRESENTATION', 'K36'),
    ('PRESENTATION', 'T35'),
    ('PRESENTATION', 'AA35'),
    ('PRESENTATION', 'F37'),
    ('PRESENTATION', 'M37'),
    ('PRESENTATION', 'T37'),
    ('PRESENTATION', 'AA37'),
    ('PRESENTATION', 'T39'),
    ('PRESENTATION', 'AA39'),
    ('PRESENTATION', 'F41'),
    ('PRESENTATION', 'M41'),
    ('PRESENTATION', 'T41'),
    ('PRESENTATION', 'AA41'),
    ('PRESENTATION', 'F42'),
    ('PRESENTATION', 'M42'),
    ('PRESENTATION', 'T42'),
    ('PRESENTATION', 'AA42'),
    ('PRESENTATION', 'F43'),
    ('PRESENTATION', 'M43'),
    ('PRESENTATION', 'T43'),
    ('PRESENTATION', 'AA43'),
    ('Segment Revenue Model', 'K24'),
    ('Segment Revenue Model', 'U8'),
    ('Segment Revenue Model', 'P24'),
    ('Segment Revenue Model', 'K25'),
    ('Segment Revenue Model', 'U9'),
    ('Segment Revenue Model', 'P25'),
    ('Segment Revenue Model', 'K26'),
    ('Segment Revenue Model', 'U10'),
    ('Segment Revenue Model', 'P26'),
    ('Segment Revenue Model', 'K27'),
    ('Segment Revenue Model', 'U11'),
    ('Segment Revenue Model', 'P27'),
    ('Segment Revenue Model', 'K28'),
    ('Segment Revenue Model', 'P28'),
    ('Segment Revenue Model', 'T12'),
    ('Segment Revenue Model', 'V12'),
    ('Segment Revenue Model', 'K29'),
    ('Segment Revenue Model', 'U13'),
    ('Segment Revenue Model', 'P29'),
    ('Segment Revenue Model', 'F15'),
    ('Segment Revenue Model', 'K15'),
    ('Segment Revenue Model', 'K30'),
    ('Segment Revenue Model', 'U14'),
    ('Segment Revenue Model', 'P15'),
    ('Segment Revenue Model', 'P30'),
    ('Segment Revenue Model', 'B17'),
    ('Segment Revenue Model', 'C17'),
    ('Segment Revenue Model', 'D17'),
    ('Segment Revenue Model', 'E17'),
    ('Segment Revenue Model', 'G17'),
    ('Segment Revenue Model', 'G31'),
    ('Segment Revenue Model', 'H17'),
    ('Segment Revenue Model', 'H31'),
    ('Segment Revenue Model', 'I17'),
    ('Segment Revenue Model', 'I31'),
    ('Segment Revenue Model', 'J17'),
    ('Segment Revenue Model', 'J31'),
    ('Segment Revenue Model', 'L17'),
    ('Segment Revenue Model', 'L31'),
    ('Segment Revenue Model', 'M17'),
    ('Segment Revenue Model', 'M31'),
    ('Segment Revenue Model', 'N17'),
    ('Segment Revenue Model', 'N31'),
    ('Segment Revenue Model', 'O17'),
    ('Segment Revenue Model', 'O31'),
    ('Segment Revenue Model', 'Q17'),
    ('Segment Revenue Model', 'Q31'),
    ('Segment Revenue Model', 'R17'),
    ('Segment Revenue Model', 'R31'),
    ('Segment Revenue Model', 'S17'),
    ('Segment Revenue Model', 'S31'),
    ('Segment Revenue Model', 'K32'),
    ('Segment Revenue Model', 'P32'),
    ('Segment Revenue Model', 'V16'),
    ('Segment Revenue Model', 'F58'),
    ('Segment Revenue Model', 'K58'),
    ('Segment Revenue Model', 'P58'),
    ('Segment Revenue Model', 'U58'),
    ('Segment Revenue Model', 'B61'),
    ('Segment Revenue Model', 'C61'),
    ('Segment Revenue Model', 'D61'),
    ('Segment Revenue Model', 'E61'),
    ('Segment Revenue Model', 'G61'),
    ('Segment Revenue Model', 'H61'),
    ('Segment Revenue Model', 'I61'),
    ('Segment Revenue Model', 'J61'),
    ('Segment Revenue Model', 'M61'),
    ('Segment Revenue Model', 'N61'),
    ('Segment Revenue Model', 'O61'),
    ('Segment Revenue Model', 'Q61'),
    ('Segment Revenue Model', 'R61'),
    ('Segment Revenue Model', 'S61'),
    ('Segment Revenue Model', 'T61'),
    ('Segment Revenue Model', 'V61'),
    ('Segment Revenue Model', 'W61'),
    ('Segment Revenue Model', 'X61'),
    ('Segment Revenue Model', 'Y61'),
    ('Segment Revenue Model', 'F72'),
    ('Segment Revenue Model', 'K72'),
    ('Segment Revenue Model', 'P72'),
    ('Segment Revenue Model', 'B74'),
    ('Segment Revenue Model', 'C74'),
    ('Segment Revenue Model', 'D74'),
    ('Segment Revenue Model', 'E74'),
    ('Segment Revenue Model', 'G74'),
    ('Segment Revenue Model', 'H74'),
    ('Segment Revenue Model', 'I74'),
    ('Segment Revenue Model', 'J74'),
    ('Segment Revenue Model', 'L74'),
    ('Segment Revenue Model', 'M74'),
    ('Segment Revenue Model', 'N74'),
    ('Segment Revenue Model', 'O74'),
    ('Segment Revenue Model', 'Q74'),
    ('Segment Revenue Model', 'R74'),
    ('Segment Revenue Model', 'S74'),
    ('Segment Revenue Model', 'U72'),
    ('Segment Revenue Model', 'T74'),
    ('INCOME STATEMENT', 'F6'),
    ('INCOME STATEMENT', 'D7'),
    ('INCOME STATEMENT', 'K8'),
    ('INCOME STATEMENT', 'M6'),
    ('INCOME STATEMENT', 'T6'),
    ('INCOME STATEMENT', 'R7'),
    ('INCOME STATEMENT', 'R8'),
    ('INCOME STATEMENT', 'AA6'),
    ('INCOME STATEMENT', 'Y8'),
    ('Ratio Analysis', 'I21'),
    ('INCOME STATEMENT', 'B21'),
    ('INCOME STATEMENT', 'C21'),
    ('INCOME STATEMENT', 'E21'),
    ('INCOME STATEMENT', 'G21'),
    ('INCOME STATEMENT', 'I21'),
    ('INCOME STATEMENT', 'J21'),
    ('INCOME STATEMENT', 'L21'),
    ('INCOME STATEMENT', 'N21'),
    ('INCOME STATEMENT', 'P21'),
    ('INCOME STATEMENT', 'Q21'),
    ('INCOME STATEMENT', 'S21'),
    ('INCOME STATEMENT', 'U21'),
    ('INCOME STATEMENT', 'W21'),
    ('INCOME STATEMENT', 'X21'),
    ('INCOME STATEMENT', 'Z21'),
    ('INCOME STATEMENT', 'AB21'),
    ('INCOME STATEMENT', 'F11'),
    ('INCOME STATEMENT', 'D12'),
    ('INCOME STATEMENT', 'M11'),
    ('INCOME STATEMENT', 'K12'),
    ('INCOME STATEMENT', 'T11'),
    ('INCOME STATEMENT', 'R12'),
    ('INCOME STATEMENT', 'AA11'),
    ('INCOME STATEMENT', 'Y12'),
    ('INCOME STATEMENT', 'F13'),
    ('INCOME STATEMENT', 'D14'),
    ('INCOME STATEMENT', 'M13'),
    ('INCOME STATEMENT', 'K14'),
    ('INCOME STATEMENT', 'T13'),
    ('INCOME STATEMENT', 'R14'),
    ('INCOME STATEMENT', 'AA13'),
    ('INCOME STATEMENT', 'Y14'),
    ('INCOME STATEMENT', 'F15'),
    ('INCOME STATEMENT', 'D16'),
    ('INCOME STATEMENT', 'M15'),
    ('INCOME STATEMENT', 'K16'),
    ('INCOME STATEMENT', 'T15'),
    ('INCOME STATEMENT', 'R16'),
    ('INCOME STATEMENT', 'AA15'),
    ('INCOME STATEMENT', 'Y16'),
    ('INCOME STATEMENT', 'F17'),
    ('INCOME STATEMENT', 'D18'),
    ('INCOME STATEMENT', 'M17'),
    ('INCOME STATEMENT', 'K18'),
    ('INCOME STATEMENT', 'T17'),
    ('INCOME STATEMENT', 'R18'),
    ('INCOME STATEMENT', 'AA17'),
    ('INCOME STATEMENT', 'Y18'),
    ('INCOME STATEMENT', 'D10'),
    ('INCOME STATEMENT', 'F19'),
    ('INCOME STATEMENT', 'D20'),
    ('INCOME STATEMENT', 'M19'),
    ('INCOME STATEMENT', 'K20'),
    ('INCOME STATEMENT', 'R10'),
    ('INCOME STATEMENT', 'T19'),
    ('INCOME STATEMENT', 'R20'),
    ('INCOME STATEMENT', 'Y10'),
    ('INCOME STATEMENT', 'AA19'),
    ('INCOME STATEMENT', 'Y20'),
    ('INCOME STATEMENT', 'AD10'),
    ('INCOME STATEMENT', 'AE10'),
    ('INCOME STATEMENT', 'AF10'),
    ('INCOME STATEMENT', 'AG10'),
    ('INCOME STATEMENT', 'K25'),
    ('INCOME STATEMENT', 'F26'),
    ('INCOME STATEMENT', 'M26'),
    ('INCOME STATEMENT', 'T26'),
    ('INCOME STATEMENT', 'AA26'),
    ('INCOME STATEMENT', 'F27'),
    ('INCOME STATEMENT', 'M27'),
    ('INCOME STATEMENT', 'K28'),
    ('INCOME STATEMENT', 'T27'),
    ('INCOME STATEMENT', 'AA27'),
    ('INCOME STATEMENT', 'F29'),
    ('INCOME STATEMENT', 'M29'),
    ('INCOME STATEMENT', 'T29'),
    ('INCOME STATEMENT', 'AA29'),
    ('INCOME STATEMENT', 'T31'),
    ('INCOME STATEMENT', 'AA31'),
    ('INCOME STATEMENT', 'F33'),
    ('INCOME STATEMENT', 'M33'),
    ('INCOME STATEMENT', 'T33'),
    ('INCOME STATEMENT', 'AA33'),
    ('INCOME STATEMENT', 'F34'),
    ('INCOME STATEMENT', 'M34'),
    ('INCOME STATEMENT', 'T34'),
    ('INCOME STATEMENT', 'AA34'),
    ('INCOME STATEMENT', 'F35'),
    ('INCOME STATEMENT', 'M35'),
    ('INCOME STATEMENT', 'T35'),
    ('INCOME STATEMENT', 'AA35'),
    ('INCOME STATEMENT', 'D47'),
    ('INCOME STATEMENT', 'C48'),
    ('INCOME STATEMENT', 'U47'),
    ('INCOME STATEMENT', 'T48'),
    ('CASH FOW STATEMENT', 'I10'),
    ('Valuation', 'B27'),
    ('Valuation', 'C27'),
    ('Valuation', 'D27'),
    ('Valuation', 'E27'),
    ('Valuation', 'F27'),
    ('Valuation', 'G27'),
    ('Valuation', 'D25'),
    ('Valuation', 'E25'),
    ('Valuation', 'F25'),
    ('Valuation', 'G25'),
    ('BALANCESHEET', 'B8'),
    ('BALANCESHEET', 'C8'),
    ('BALANCESHEET', 'D8'),
    ('BALANCESHEET', 'E8'),
    ('BALANCESHEET', 'F8'),
    ('BALANCESHEET', 'G8'),
    ('BALANCESHEET', 'H8'),
    ('BALANCESHEET', 'I8'),
    ('BALANCESHEET', 'J8'),
    ('BALANCESHEET', 'B24'),
    ('CASH FOW STATEMENT', 'B21'),
    ('BALANCESHEET', 'C24'),
    ('CASH FOW STATEMENT', 'C21'),
    ('BALANCESHEET', 'D24'),
    ('CASH FOW STATEMENT', 'D21'),
    ('BALANCESHEET', 'E24'),
    ('Valuation', 'G10'),
    ('CASH FOW STATEMENT', 'E21'),
    ('BALANCESHEET', 'F24'),
    ('CASH FOW STATEMENT', 'F21'),
    ('BALANCESHEET', 'G24'),
    ('Ratio Analysis', 'G15'),
    ('CASH FOW STATEMENT', 'G21'),
    ('BALANCESHEET', 'H24'),
    ('Ratio Analysis', 'H15'),
    ('CASH FOW STATEMENT', 'H21'),
    ('BALANCESHEET', 'I24'),
    ('Ratio Analysis', 'I15'),
    ('CASH FOW STATEMENT', 'I21'),
    ('BALANCESHEET', 'J24'),
    ('Ratio Analysis', 'J15'),
    ('BALANCESHEET', 'B27'),
    ('BALANCESHEET', 'C27'),
    ('BALANCESHEET', 'D27'),
    ('BALANCESHEET', 'E27'),
    ('BALANCESHEET', 'F27'),
    ('BALANCESHEET', 'G27'),
    ('BALANCESHEET', 'H27'),
    ('BALANCESHEET', 'I27'),
    ('BALANCESHEET', 'J27'),
    ('Ratio Analysis', 'F9'),
    ('Ratio Analysis', 'G9'),
    ('Ratio Analysis', 'H9'),
    ('Ratio Analysis', 'I9'),
    ('Ratio Analysis', 'J9'),
    ('BALANCESHEET', 'B43'),
    ('BALANCESHEET', 'B58'),
    ('BALANCESHEET', 'C43'),
    ('BALANCESHEET', 'C58'),
    ('BALANCESHEET', 'D43'),
    ('BALANCESHEET', 'D58'),
    ('BALANCESHEET', 'E43'),
    ('BALANCESHEET', 'E58'),
    ('Debt Schedule', 'B21'),
    ('Debt Schedule', 'C21'),
    ('Debt Schedule', 'D21'),
    ('Debt Schedule', 'E21'),
    ('Debt Schedule', 'F21'),
    ('Debt Schedule', 'G21'),
    ('Debt Schedule', 'H21'),
    ('Debt Schedule', 'I21'),
    ('Valuation', 'G54'),
    ('Valuation', 'E37'),
    ('Valuation', 'D40'),
    ('Valuation', 'D45'),
    ('Valuation', 'E40'),
    ('Valuation', 'E45'),
    ('Valuation', 'F40'),
    ('Valuation', 'G40'),
    ('PRESENTATION', 'G65'),
    ('PRESENTATION', 'H65'),
    ('PRESENTATION', 'I65'),
    ('PRESENTATION', 'J65'),
    ('PRESENTATION', 'I58'),
    ('PRESENTATION', 'J58'),
    ('PRESENTATION', 'K17'),
    ('PRESENTATION', 'P17'),
    ('PRESENTATION', 'U15'),
    ('PRESENTATION', 'T17'),
    ('PRESENTATION', 'F17'),
    ('PRESENTATION', 'O25'),
    ('PRESENTATION', 'AC25'),
    ('PRESENTATION', 'B33'),
    ('PRESENTATION', 'B36'),
    ('PRESENTATION', 'C33'),
    ('PRESENTATION', 'C36'),
    ('PRESENTATION', 'E33'),
    ('PRESENTATION', 'E36'),
    ('PRESENTATION', 'G33'),
    ('PRESENTATION', 'G36'),
    ('PRESENTATION', 'I33'),
    ('PRESENTATION', 'I36'),
    ('PRESENTATION', 'J33'),
    ('PRESENTATION', 'J36'),
    ('PRESENTATION', 'L33'),
    ('PRESENTATION', 'L36'),
    ('PRESENTATION', 'N33'),
    ('PRESENTATION', 'N36'),
    ('PRESENTATION', 'P33'),
    ('PRESENTATION', 'P36'),
    ('PRESENTATION', 'Q33'),
    ('PRESENTATION', 'Q36'),
    ('PRESENTATION', 'S33'),
    ('PRESENTATION', 'S36'),
    ('PRESENTATION', 'U33'),
    ('PRESENTATION', 'U36'),
    ('PRESENTATION', 'W33'),
    ('PRESENTATION', 'W36'),
    ('PRESENTATION', 'X33'),
    ('PRESENTATION', 'X36'),
    ('PRESENTATION', 'Z33'),
    ('PRESENTATION', 'Z36'),
    ('PRESENTATION', 'AB33'),
    ('PRESENTATION', 'AB36'),
    ('PRESENTATION', 'AD33'),
    ('PRESENTATION', 'AD36'),
    ('PRESENTATION', 'AE33'),
    ('PRESENTATION', 'AE36'),
    ('PRESENTATION', 'AF33'),
    ('PRESENTATION', 'AF36'),
    ('PRESENTATION', 'AG33'),
    ('PRESENTATION', 'AG36'),
    ('PRESENTATION', 'O27'),
    ('PRESENTATION', 'AC27'),
    ('PRESENTATION', 'AC28'),
    ('PRESENTATION', 'H29'),
    ('PRESENTATION', 'O29'),
    ('PRESENTATION', 'V29'),
    ('PRESENTATION', 'AC29'),
    ('PRESENTATION', 'O30'),
    ('PRESENTATION', 'AC30'),
    ('PRESENTATION', 'D32'),
    ('PRESENTATION', 'F26'),
    ('PRESENTATION', 'H31'),
    ('PRESENTATION', 'M26'),
    ('PRESENTATION', 'R32'),
    ('PRESENTATION', 'T26'),
    ('PRESENTATION', 'V31'),
    ('PRESENTATION', 'Y32'),
    ('PRESENTATION', 'AA26'),
    ('PRESENTATION', 'AC31'),
    ('PRESENTATION', 'H34'),
    ('PRESENTATION', 'O34'),
    ('PRESENTATION', 'V34'),
    ('PRESENTATION', 'AC34'),
    ('PRESENTATION', 'H35'),
    ('PRESENTATION', 'O35'),
    ('PRESENTATION', 'K38'),
    ('PRESENTATION', 'V35'),
    ('PRESENTATION', 'H37'),
    ('PRESENTATION', 'O37'),
    ('PRESENTATION', 'AC37'),
    ('PRESENTATION', 'AC41'),
    ('PRESENTATION', 'H42'),
    ('PRESENTATION', 'AC42'),
    ('PRESENTATION', 'H43'),
    ('Segment Revenue Model', 'T8'),
    ('Segment Revenue Model', 'V8'),
    ('Segment Revenue Model', 'T9'),
    ('Segment Revenue Model', 'V9'),
    ('Segment Revenue Model', 'T10'),
    ('Segment Revenue Model', 'V10'),
    ('Segment Revenue Model', 'T11'),
    ('Segment Revenue Model', 'V11'),
    ('Segment Revenue Model', 'T28'),
    ('Segment Revenue Model', 'W12'),
    ('Segment Revenue Model', 'T13'),
    ('Segment Revenue Model', 'V13'),
    ('Segment Revenue Model', 'K17'),
    ('Segment Revenue Model', 'K31'),
    ('Segment Revenue Model', 'T14'),
    ('Segment Revenue Model', 'V14'),
    ('Segment Revenue Model', 'P17'),
    ('Segment Revenue Model', 'P31'),
    ('Segment Revenue Model', 'F17'),
    ('Segment Revenue Model', 'G33'),
    ('Segment Revenue Model', 'G37'),
    ('Segment Revenue Model', 'G38'),
    ('Segment Revenue Model', 'G39'),
    ('Segment Revenue Model', 'G40'),
    ('Segment Revenue Model', 'G41'),
    ('Segment Revenue Model', 'G42'),
    ('Segment Revenue Model', 'G43'),
    ('Segment Revenue Model', 'G44'),
    ('Segment Revenue Model', 'G45'),
    ('Segment Revenue Model', 'G46'),
    ('Segment Revenue Model', 'H33'),
    ('Segment Revenue Model', 'H37'),
    ('Segment Revenue Model', 'H38'),
    ('Segment Revenue Model', 'H39'),
    ('Segment Revenue Model', 'H40'),
    ('Segment Revenue Model', 'H41'),
    ('Segment Revenue Model', 'H42'),
    ('Segment Revenue Model', 'H43'),
    ('Segment Revenue Model', 'H44'),
    ('Segment Revenue Model', 'H45'),
    ('Segment Revenue Model', 'H46'),
    ('Segment Revenue Model', 'I33'),
    ('Segment Revenue Model', 'I37'),
    ('Segment Revenue Model', 'I38'),
    ('Segment Revenue Model', 'I39'),
    ('Segment Revenue Model', 'I40'),
    ('Segment Revenue Model', 'I41'),
    ('Segment Revenue Model', 'I42'),
    ('Segment Revenue Model', 'I43'),
    ('Segment Revenue Model', 'I44'),
    ('Segment Revenue Model', 'I45'),
    ('Segment Revenue Model', 'I46'),
    ('Segment Revenue Model', 'J33'),
    ('Segment Revenue Model', 'J37'),
    ('Segment Revenue Model', 'J38'),
    ('Segment Revenue Model', 'J39'),
    ('Segment Revenue Model', 'J40'),
    ('Segment Revenue Model', 'J41'),
    ('Segment Revenue Model', 'J42'),
    ('Segment Revenue Model', 'J43'),
    ('Segment Revenue Model', 'J44'),
    ('Segment Revenue Model', 'J45'),
    ('Segment Revenue Model', 'J46'),
    ('Segment Revenue Model', 'L33'),
    ('Segment Revenue Model', 'L37'),
    ('Segment Revenue Model', 'L38'),
    ('Segment Revenue Model', 'L39'),
    ('Segment Revenue Model', 'L40'),
    ('Segment Revenue Model', 'L41'),
    ('Segment Revenue Model', 'L42'),
    ('Segment Revenue Model', 'L43'),
    ('Segment Revenue Model', 'L44'),
    ('Segment Revenue Model', 'L45'),
    ('Segment Revenue Model', 'L46'),
    ('Segment Revenue Model', 'M33'),
    ('Segment Revenue Model', 'M37'),
    ('Segment Revenue Model', 'M38'),
    ('Segment Revenue Model', 'M39'),
    ('Segment Revenue Model', 'M40'),
    ('Segment Revenue Model', 'M41'),
    ('Segment Revenue Model', 'M42'),
    ('Segment Revenue Model', 'M43'),
    ('Segment Revenue Model', 'M44'),
    ('Segment Revenue Model', 'M45'),
    ('Segment Revenue Model', 'M46'),
    ('Segment Revenue Model', 'N33'),
    ('Segment Revenue Model', 'N37'),
    ('Segment Revenue Model', 'N38'),
    ('Segment Revenue Model', 'N39'),
    ('Segment Revenue Model', 'N40'),
    ('Segment Revenue Model', 'N41'),
    ('Segment Revenue Model', 'N42'),
    ('Segment Revenue Model', 'N43'),
    ('Segment Revenue Model', 'N44'),
    ('Segment Revenue Model', 'N45'),
    ('Segment Revenue Model', 'N46'),
    ('Segment Revenue Model', 'O33'),
    ('Segment Revenue Model', 'O37'),
    ('Segment Revenue Model', 'O38'),
    ('Segment Revenue Model', 'O39'),
    ('Segment Revenue Model', 'O40'),
    ('Segment Revenue Model', 'O41'),
    ('Segment Revenue Model', 'O42'),
    ('Segment Revenue Model', 'O43'),
    ('Segment Revenue Model', 'O44'),
    ('Segment Revenue Model', 'O45'),
    ('Segment Revenue Model', 'O46'),
    ('Segment Revenue Model', 'Q33'),
    ('Segment Revenue Model', 'Q37'),
    ('Segment Revenue Model', 'Q38'),
    ('Segment Revenue Model', 'Q39'),
    ('Segment Revenue Model', 'Q40'),
    ('Segment Revenue Model', 'Q41'),
    ('Segment Revenue Model', 'Q42'),
    ('Segment Revenue Model', 'Q43'),
    ('Segment Revenue Model', 'Q44'),
    ('Segment Revenue Model', 'Q45'),
    ('Segment Revenue Model', 'Q46'),
    ('Segment Revenue Model', 'R33'),
    ('Segment Revenue Model', 'R37'),
    ('Segment Revenue Model', 'R38'),
    ('Segment Revenue Model', 'R39'),
    ('Segment Revenue Model', 'R40'),
    ('Segment Revenue Model', 'R41'),
    ('Segment Revenue Model', 'R42'),
    ('Segment Revenue Model', 'R43'),
    ('Segment Revenue Model', 'R44'),
    ('Segment Revenue Model', 'R45'),
    ('Segment Revenue Model', 'R46'),
    ('Segment Revenue Model', 'S33'),
    ('Segment Revenue Model', 'S37'),
    ('Segment Revenue Model', 'S38'),
    ('Segment Revenue Model', 'S39'),
    ('Segment Revenue Model', 'S40'),
    ('Segment Revenue Model', 'S41'),
    ('Segment Revenue Model', 'S42'),
    ('Segment Revenue Model', 'S43'),
    ('Segment Revenue Model', 'S44'),
    ('Segment Revenue Model', 'S45'),
    ('Segment Revenue Model', 'S46'),
    ('Segment Revenue Model', 'W16'),
    ('Segment Revenue Model', 'F61'),
    ('Segment Revenue Model', 'P61'),
    ('Segment Revenue Model', 'U61'),
    ('Segment Revenue Model', 'K61'),
    ('Segment Revenue Model', 'K74'),
    ('Segment Revenue Model', 'P74'),
    ('Segment Revenue Model', 'F74'),
    ('Segment Revenue Model', 'U74'),
    ('INCOME STATEMENT', 'F7'),
    ('INCOME STATEMENT', 'O6'),
    ('INCOME STATEMENT', 'M8'),
    ('INCOME STATEMENT', 'T7'),
    ('INCOME STATEMENT', 'T8'),
    ('INCOME STATEMENT', 'AC6'),
    ('INCOME STATEMENT', 'AA8'),
    ('INCOME STATEMENT', 'B25'),
    ('INCOME STATEMENT', 'B28'),
    ('INCOME STATEMENT', 'C24'),
    ('INCOME STATEMENT', 'C25'),
    ('INCOME STATEMENT', 'C28'),
    ('INCOME STATEMENT', 'E24'),
    ('INCOME STATEMENT', 'E25'),
    ('INCOME STATEMENT', 'E28'),
    ('INCOME STATEMENT', 'G24'),
    ('INCOME STATEMENT', 'G25'),
    ('INCOME STATEMENT', 'G28'),
    ('INCOME STATEMENT', 'I23'),
    ('INCOME STATEMENT', 'I24'),
    ('INCOME STATEMENT', 'I25'),
    ('INCOME STATEMENT', 'I28'),
    ('INCOME STATEMENT', 'J23'),
    ('INCOME STATEMENT', 'J24'),
    ('INCOME STATEMENT', 'J25'),
    ('INCOME STATEMENT', 'J28'),
    ('INCOME STATEMENT', 'L23'),
    ('INCOME STATEMENT', 'L24'),
    ('INCOME STATEMENT', 'L25'),
    ('INCOME STATEMENT', 'L28'),
    ('INCOME STATEMENT', 'N23'),
    ('INCOME STATEMENT', 'N24'),
    ('INCOME STATEMENT', 'N25'),
    ('INCOME STATEMENT', 'N28'),
    ('INCOME STATEMENT', 'P23'),
    ('INCOME STATEMENT', 'P24'),
    ('INCOME STATEMENT', 'P25'),
    ('INCOME STATEMENT', 'P28'),
    ('INCOME STATEMENT', 'Q23'),
    ('INCOME STATEMENT', 'Q24'),
    ('INCOME STATEMENT', 'Q25'),
    ('INCOME STATEMENT', 'Q28'),
    ('INCOME STATEMENT', 'S23'),
    ('INCOME STATEMENT', 'S24'),
    ('INCOME STATEMENT', 'S25'),
    ('INCOME STATEMENT', 'S28'),
    ('INCOME STATEMENT', 'U23'),
    ('INCOME STATEMENT', 'U24'),
    ('INCOME STATEMENT', 'U25'),
    ('INCOME STATEMENT', 'U28'),
    ('INCOME STATEMENT', 'W23'),
    ('INCOME STATEMENT', 'W24'),
    ('INCOME STATEMENT', 'W25'),
    ('INCOME STATEMENT', 'W28'),
    ('INCOME STATEMENT', 'X23'),
    ('INCOME STATEMENT', 'X24'),
    ('INCOME STATEMENT', 'X25'),
    ('INCOME STATEMENT', 'X28'),
    ('INCOME STATEMENT', 'Z23'),
    ('INCOME STATEMENT', 'Z24'),
    ('INCOME STATEMENT', 'Z25'),
    ('INCOME STATEMENT', 'Z28'),
    ('INCOME STATEMENT', 'AB23'),
    ('INCOME STATEMENT', 'AB24'),
    ('INCOME STATEMENT', 'AB25'),
    ('INCOME STATEMENT', 'AB28'),
    ('INCOME STATEMENT', 'F12'),
    ('INCOME STATEMENT', 'O11'),
    ('INCOME STATEMENT', 'M12'),
    ('INCOME STATEMENT', 'T12'),
    ('INCOME STATEMENT', 'AC11'),
    ('INCOME STATEMENT', 'AA12'),
    ('INCOME STATEMENT', 'F14'),
    ('INCOME STATEMENT', 'M14'),
    ('INCOME STATEMENT', 'T14'),
    ('INCOME STATEMENT', 'AC13'),
    ('INCOME STATEMENT', 'AA14'),
    ('INCOME STATEMENT', 'H15'),
    ('INCOME STATEMENT', 'F16'),
    ('INCOME STATEMENT', 'O15'),
    ('INCOME STATEMENT', 'M16'),
    ('INCOME STATEMENT', 'V15'),
    ('INCOME STATEMENT', 'T16'),
    ('INCOME STATEMENT', 'AC15'),
    ('INCOME STATEMENT', 'AA16'),
    ('INCOME STATEMENT', 'F18'),
    ('INCOME STATEMENT', 'O17'),
    ('INCOME STATEMENT', 'M18'),
    ('INCOME STATEMENT', 'T18'),
    ('INCOME STATEMENT', 'AC17'),
    ('INCOME STATEMENT', 'AA18'),
    ('INCOME STATEMENT', 'D21'),
    ('INCOME STATEMENT', 'F10'),
    ('INCOME STATEMENT', 'H19'),
    ('INCOME STATEMENT', 'F20'),
    ('INCOME STATEMENT', 'M10'),
    ('INCOME STATEMENT', 'M20'),
    ('INCOME STATEMENT', 'R21'),
    ('INCOME STATEMENT', 'T10'),
    ('INCOME STATEMENT', 'V19'),
    ('INCOME STATEMENT', 'T20'),
    ('INCOME STATEMENT', 'Y21'),
    ('INCOME STATEMENT', 'AA10'),
    ('INCOME STATEMENT', 'AC19'),
    ('INCOME STATEMENT', 'AA20'),
    ('PRESENTATION', 'G64'),
    ('INCOME STATEMENT', 'AD21'),
    ('PRESENTATION', 'H64'),
    ('INCOME STATEMENT', 'AE21'),
    ('PRESENTATION', 'I64'),
    ('INCOME STATEMENT', 'AF21'),
    ('PRESENTATION', 'J64'),
    ('INCOME STATEMENT', 'AG21'),
    ('INCOME STATEMENT', 'H26'),
    ('INCOME STATEMENT', 'O26'),
    ('INCOME STATEMENT', 'V26'),
    ('INCOME STATEMENT', 'AC26'),
    ('INCOME STATEMENT', 'H27'),
    ('INCOME STATEMENT', 'O27'),
    ('INCOME STATEMENT', 'K30'),
    ('INCOME STATEMENT', 'V27'),
    ('INCOME STATEMENT', 'H29'),
    ('INCOME STATEMENT', 'O29'),
    ('INCOME STATEMENT', 'AC29'),
    ('INCOME STATEMENT', 'AC33'),
    ('INCOME STATEMENT', 'H34'),
    ('INCOME STATEMENT', 'AC34'),
    ('INCOME STATEMENT', 'H35'),
    ('INCOME STATEMENT', 'E47'),
    ('INCOME STATEMENT', 'D48'),
    ('INCOME STATEMENT', 'V47'),
    ('INCOME STATEMENT', 'U48'),
    ('CASH FOW STATEMENT', 'I12'),
    ('Valuation', 'B46'),
    ('Valuation', 'D46'),
    ('Valuation', 'E46'),
    ('Valuation', 'H25'),
    ('Ratio Analysis', 'B8'),
    ('Ratio Analysis', 'C8'),
    ('Ratio Analysis', 'D8'),
    ('Ratio Analysis', 'E8'),
    ('Ratio Analysis', 'F8'),
    ('Ratio Analysis', 'G8'),
    ('Ratio Analysis', 'H8'),
    ('Ratio Analysis', 'I8'),
    ('Ratio Analysis', 'J8'),
    ('Ratio Analysis', 'B7'),
    ('Ratio Analysis', 'C7'),
    ('Ratio Analysis', 'D7'),
    ('Ratio Analysis', 'E7'),
    ('Valuation', 'G12'),
    ('Valuation', 'G53'),
    ('Ratio Analysis', 'F7'),
    ('Ratio Analysis', 'G7'),
    ('Ratio Analysis', 'H7'),
    ('Ratio Analysis', 'I7'),
    ('Ratio Analysis', 'J7'),
    ('Ratio Analysis', 'B17'),
    ('Ratio Analysis', 'C10'),
    ('Ratio Analysis', 'C17'),
    ('Ratio Analysis', 'D17'),
    ('Ratio Analysis', 'E10'),
    ('Ratio Analysis', 'E17'),
    ('BALANCESHEET', 'F60'),
    ('Ratio Analysis', 'F17'),
    ('BALANCESHEET', 'G60'),
    ('Ratio Analysis', 'G10'),
    ('Ratio Analysis', 'G17'),
    ('BALANCESHEET', 'H60'),
    ('Ratio Analysis', 'H10'),
    ('Ratio Analysis', 'H17'),
    ('BALANCESHEET', 'I60'),
    ('Ratio Analysis', 'I10'),
    ('Ratio Analysis', 'I17'),
    ('BALANCESHEET', 'J60'),
    ('Ratio Analysis', 'J10'),
    ('Ratio Analysis', 'J17'),
    ('Ratio Analysis', 'B9'),
    ('BALANCESHEET', 'B60'),
    ('CASH FOW STATEMENT', 'B20'),
    ('Ratio Analysis', 'C9'),
    ('BALANCESHEET', 'C60'),
    ('Ratio Analysis', 'C24'),
    ('CASH FOW STATEMENT', 'C20'),
    ('Ratio Analysis', 'D9'),
    ('BALANCESHEET', 'D60'),
    ('CASH FOW STATEMENT', 'D20'),
    ('CASH FOW STATEMENT', 'E20'),
    ('Ratio Analysis', 'E9'),
    ('BALANCESHEET', 'E60'),
    ('Ratio Analysis', 'E24'),
    ('Valuation', 'F37'),
    ('Valuation', 'F45'),
    ('PRESENTATION', 'U17'),
    ('PRESENTATION', 'B38'),
    ('PRESENTATION', 'C38'),
    ('PRESENTATION', 'E38'),
    ('PRESENTATION', 'G38'),
    ('PRESENTATION', 'I38'),
    ('PRESENTATION', 'J38'),
    ('PRESENTATION', 'L38'),
    ('PRESENTATION', 'N38'),
    ('PRESENTATION', 'P38'),
    ('PRESENTATION', 'Q38'),
    ('PRESENTATION', 'S38'),
    ('PRESENTATION', 'U38'),
    ('PRESENTATION', 'W38'),
    ('PRESENTATION', 'X38'),
    ('PRESENTATION', 'Z38'),
    ('PRESENTATION', 'AB38'),
    ('PRESENTATION', 'AD38'),
    ('PRESENTATION', 'AE38'),
    ('PRESENTATION', 'AF38'),
    ('PRESENTATION', 'AG38'),
    ('PRESENTATION', 'H26'),
    ('PRESENTATION', 'O26'),
    ('PRESENTATION', 'D33'),
    ('PRESENTATION', 'D36'),
    ('PRESENTATION', 'F32'),
    ('PRESENTATION', 'M32'),
    ('PRESENTATION', 'R33'),
    ('PRESENTATION', 'R36'),
    ('PRESENTATION', 'T32'),
    ('PRESENTATION', 'V26'),
    ('PRESENTATION', 'Y33'),
    ('PRESENTATION', 'Y36'),
    ('PRESENTATION', 'AA32'),
    ('PRESENTATION', 'AC26'),
    ('PRESENTATION', 'K40'),
    ('Segment Revenue Model', 'T24'),
    ('Segment Revenue Model', 'W8'),
    ('Segment Revenue Model', 'T25'),
    ('Segment Revenue Model', 'W9'),
    ('Segment Revenue Model', 'T26'),
    ('Segment Revenue Model', 'W10'),
    ('Segment Revenue Model', 'T27'),
    ('Segment Revenue Model', 'W11'),
    ('Segment Revenue Model', 'X12'),
    ('Segment Revenue Model', 'T29'),
    ('Segment Revenue Model', 'W13'),
    ('Segment Revenue Model', 'K37'),
    ('Segment Revenue Model', 'K38'),
    ('Segment Revenue Model', 'K39'),
    ('Segment Revenue Model', 'K40'),
    ('Segment Revenue Model', 'K41'),
    ('Segment Revenue Model', 'K42'),
    ('Segment Revenue Model', 'K43'),
    ('Segment Revenue Model', 'K44'),
    ('Segment Revenue Model', 'K45'),
    ('Segment Revenue Model', 'K46'),
    ('Segment Revenue Model', 'T15'),
    ('Segment Revenue Model', 'T30'),
    ('Segment Revenue Model', 'W14'),
    ('Segment Revenue Model', 'V15'),
    ('Segment Revenue Model', 'P33'),
    ('Segment Revenue Model', 'P37'),
    ('Segment Revenue Model', 'P38'),
    ('Segment Revenue Model', 'P39'),
    ('Segment Revenue Model', 'P40'),
    ('Segment Revenue Model', 'P41'),
    ('Segment Revenue Model', 'P42'),
    ('Segment Revenue Model', 'P43'),
    ('Segment Revenue Model', 'P44'),
    ('Segment Revenue Model', 'P45'),
    ('Segment Revenue Model', 'P46'),
    ('Segment Revenue Model', 'K33'),
    ('Segment Revenue Model', 'X16'),
    ('INCOME STATEMENT', 'H7'),
    ('INCOME STATEMENT', 'I7'),
    ('INCOME STATEMENT', 'J7'),
    ('INCOME STATEMENT', 'K7'),
    ('INCOME STATEMENT', 'L7'),
    ('INCOME STATEMENT', 'M7'),
    ('INCOME STATEMENT', 'N7'),
    ('INCOME STATEMENT', 'O8'),
    ('INCOME STATEMENT', 'V8'),
    ('INCOME STATEMENT', 'O14'),
    ('INCOME STATEMENT', 'O20'),
    ('Ratio Analysis', 'D10'),
    ('Ratio Analysis', 'D21'),
    ('Ratio Analysis', 'E21'),
    ('Ratio Analysis', 'D23'),
    ('Ratio Analysis', 'D24'),
    ('INCOME STATEMENT', 'V7'),
    ('COMPANY OVERVIEW', 'F28'),
    ('INCOME STATEMENT', 'W7'),
    ('INCOME STATEMENT', 'X7'),
    ('INCOME STATEMENT', 'Y7'),
    ('INCOME STATEMENT', 'Z7'),
    ('INCOME STATEMENT', 'AA7'),
    ('INCOME STATEMENT', 'AB7'),
    ('INCOME STATEMENT', 'AC8'),
    ('INCOME STATEMENT', 'AD8'),
    ('Valuation', 'C19'),
    ('Ratio Analysis', 'F10'),
    ('Ratio Analysis', 'F21'),
    ('Ratio Analysis', 'F23'),
    ('Ratio Analysis', 'F24'),
    ('INCOME STATEMENT', 'B30'),
    ('INCOME STATEMENT', 'C30'),
    ('INCOME STATEMENT', 'E30'),
    ('INCOME STATEMENT', 'G30'),
    ('INCOME STATEMENT', 'I30'),
    ('INCOME STATEMENT', 'J30'),
    ('INCOME STATEMENT', 'L30'),
    ('INCOME STATEMENT', 'N30'),
    ('INCOME STATEMENT', 'P30'),
    ('INCOME STATEMENT', 'Q30'),
    ('INCOME STATEMENT', 'S30'),
    ('INCOME STATEMENT', 'U30'),
    ('INCOME STATEMENT', 'W30'),
    ('INCOME STATEMENT', 'X30'),
    ('INCOME STATEMENT', 'Z30'),
    ('INCOME STATEMENT', 'AB30'),
    ('INCOME STATEMENT', 'O12'),
    ('INCOME STATEMENT', 'AC12'),
    ('INCOME STATEMENT', 'AC14'),
    ('INCOME STATEMENT', 'H10'),
    ('INCOME STATEMENT', 'H16'),
    ('INCOME STATEMENT', 'O16'),
    ('INCOME STATEMENT', 'V16'),
    ('INCOME STATEMENT', 'AC16'),
    ('INCOME STATEMENT', 'O10'),
    ('INCOME STATEMENT', 'O18'),
    ('INCOME STATEMENT', 'AC18'),
    ('INCOME STATEMENT', 'K23'),
    ('INCOME STATEMENT', 'D25'),
    ('INCOME STATEMENT', 'D28'),
    ('INCOME STATEMENT', 'F21'),
    ('INCOME STATEMENT', 'H20'),
    ('INCOME STATEMENT', 'M21'),
    ('INCOME STATEMENT', 'R23'),
    ('INCOME STATEMENT', 'R25'),
    ('INCOME STATEMENT', 'R28'),
    ('INCOME STATEMENT', 'T21'),
    ('INCOME STATEMENT', 'V10'),
    ('INCOME STATEMENT', 'V20'),
    ('INCOME STATEMENT', 'Y23'),
    ('INCOME STATEMENT', 'Y25'),
    ('INCOME STATEMENT', 'Y28'),
    ('INCOME STATEMENT', 'AA21'),
    ('INCOME STATEMENT', 'AC10'),
    ('INCOME STATEMENT', 'AC20'),
    ('COMPANY OVERVIEW', 'G22'),
    ('INCOME STATEMENT', 'AD25'),
    ('INCOME STATEMENT', 'AD28'),
    ('Valuation', 'D20'),
    ('Ratio Analysis', 'G13'),
    ('Ratio Analysis', 'G19'),
    ('Ratio Analysis', 'G22'),
    ('COMPANY OVERVIEW', 'H22'),
    ('INCOME STATEMENT', 'AE23'),
    ('INCOME STATEMENT', 'AE25'),
    ('INCOME STATEMENT', 'AE28'),
    ('Valuation', 'E20'),
    ('Ratio Analysis', 'H13'),
    ('Ratio Analysis', 'H19'),
    ('Ratio Analysis', 'H22'),
    ('COMPANY OVERVIEW', 'I22'),
    ('INCOME STATEMENT', 'AF23'),
    ('INCOME STATEMENT', 'AF25'),
    ('INCOME STATEMENT', 'AF28'),
    ('Valuation', 'F20'),
    ('Ratio Analysis', 'I13'),
    ('Ratio Analysis', 'I19'),
    ('Ratio Analysis', 'I22'),
    ('COMPANY OVERVIEW', 'J22'),
    ('INCOME STATEMENT', 'AG23'),
    ('INCOME STATEMENT', 'AG25'),
    ('INCOME STATEMENT', 'AG28'),
    ('Valuation', 'G20'),
    ('Ratio Analysis', 'J13'),
    ('Ratio Analysis', 'J19'),
    ('Ratio Analysis', 'J22'),
    ('Ratio Analysis', 'C15'),
    ('Ratio Analysis', 'D15'),
    ('Ratio Analysis', 'E15'),
    ('Ratio Analysis', 'F15'),
    ('CASH FOW STATEMENT', 'B11'),
    ('CASH FOW STATEMENT', 'C11'),
    ('INCOME STATEMENT', 'K32'),
    ('CASH FOW STATEMENT', 'D11'),
    ('Valuation', 'B24'),
    ('Valuation', 'C22'),
    ('CASH FOW STATEMENT', 'B9'),
    ('CASH FOW STATEMENT', 'E9'),
    ('INCOME STATEMENT', 'F47'),
    ('INCOME STATEMENT', 'E48'),
    ('INCOME STATEMENT', 'W47'),
    ('INCOME STATEMENT', 'V48'),
    ('CASH FOW STATEMENT', 'I14'),
    ('Valuation', 'F46'),
    ('Valuation', 'I25'),
    ('Valuation', 'G13'),
    ('Valuation', 'G55'),
    ('Valuation', 'B25'),
    ('Valuation', 'C25'),
    ('Valuation', 'G37'),
    ('Valuation', 'G45'),
    ('PRESENTATION', 'B40'),
    ('PRESENTATION', 'C40'),
    ('PRESENTATION', 'E40'),
    ('PRESENTATION', 'G40'),
    ('PRESENTATION', 'I40'),
    ('PRESENTATION', 'J40'),
    ('PRESENTATION', 'L40'),
    ('PRESENTATION', 'N40'),
    ('PRESENTATION', 'P40'),
    ('PRESENTATION', 'Q40'),
    ('PRESENTATION', 'S40'),
    ('PRESENTATION', 'U40'),
    ('PRESENTATION', 'W40'),
    ('PRESENTATION', 'X40'),
    ('PRESENTATION', 'Z40'),
    ('PRESENTATION', 'AB40'),
    ('PRESENTATION', 'AD40'),
    ('PRESENTATION', 'AE40'),
    ('PRESENTATION', 'AF40'),
    ('PRESENTATION', 'AG40'),
    ('PRESENTATION', 'H32'),
    ('PRESENTATION', 'O32'),
    ('PRESENTATION', 'D38'),
    ('PRESENTATION', 'F33'),
    ('PRESENTATION', 'F36'),
    ('PRESENTATION', 'M33'),
    ('PRESENTATION', 'M36'),
    ('PRESENTATION', 'R38'),
    ('PRESENTATION', 'T33'),
    ('PRESENTATION', 'T36'),
    ('PRESENTATION', 'V32'),
    ('PRESENTATION', 'Y38'),
    ('PRESENTATION', 'AA33'),
    ('PRESENTATION', 'AA36'),
    ('PRESENTATION', 'AC32'),
    ('PRESENTATION', 'K44'),
    ('Segment Revenue Model', 'X8'),
    ('Segment Revenue Model', 'X9'),
    ('Segment Revenue Model', 'X10'),
    ('Segment Revenue Model', 'X11'),
    ('Segment Revenue Model', 'Y12'),
    ('Segment Revenue Model', 'X13'),
    ('Segment Revenue Model', 'U15'),
    ('Segment Revenue Model', 'T17'),
    ('Segment Revenue Model', 'T31'),
    ('Segment Revenue Model', 'X14'),
    ('Segment Revenue Model', 'W15'),
    ('Segment Revenue Model', 'V17'),
    ('Segment Revenue Model', 'Y16'),
    ('INCOME STATEMENT', 'O7'),
    ('INCOME STATEMENT', 'AC7'),
    ('Valuation', 'C37'),
    ('Valuation', 'D37'),
    ('Valuation', 'C45'),
    ('Valuation', 'C46'),
    ('INCOME STATEMENT', 'B32'),
    ('INCOME STATEMENT', 'C32'),
    ('INCOME STATEMENT', 'E32'),
    ('INCOME STATEMENT', 'G32'),
    ('INCOME STATEMENT', 'I32'),
    ('INCOME STATEMENT', 'J32'),
    ('INCOME STATEMENT', 'L32'),
    ('INCOME STATEMENT', 'N32'),
    ('INCOME STATEMENT', 'P32'),
    ('INCOME STATEMENT', 'Q32'),
    ('INCOME STATEMENT', 'S32'),
    ('INCOME STATEMENT', 'U32'),
    ('INCOME STATEMENT', 'W32'),
    ('INCOME STATEMENT', 'X32'),
    ('INCOME STATEMENT', 'Z32'),
    ('INCOME STATEMENT', 'AB32'),
    ('PRESENTATION', 'C64'),
    ('INCOME STATEMENT', 'H21'),
    ('PRESENTATION', 'D64'),
    ('INCOME STATEMENT', 'O21'),
    ('INCOME STATEMENT', 'D30'),
    ('INCOME STATEMENT', 'F25'),
    ('INCOME STATEMENT', 'F28'),
    ('INCOME STATEMENT', 'M23'),
    ('INCOME STATEMENT', 'M25'),
    ('INCOME STATEMENT', 'M28'),
    ('INCOME STATEMENT', 'R30'),
    ('INCOME STATEMENT', 'T23'),
    ('INCOME STATEMENT', 'T25'),
    ('INCOME STATEMENT', 'T28'),
    ('PRESENTATION', 'E64'),
    ('INCOME STATEMENT', 'V21'),
    ('INCOME STATEMENT', 'Y30'),
    ('INCOME STATEMENT', 'AA23'),
    ('INCOME STATEMENT', 'AA25'),
    ('INCOME STATEMENT', 'AA28'),
    ('PRESENTATION', 'F64'),
    ('INCOME STATEMENT', 'AC21'),
    ('INCOME STATEMENT', 'AD30'),
    ('Valuation', 'D21'),
    ('Valuation', 'D43'),
    ('PRESENTATION', 'G63'),
    ('COMPANY OVERVIEW', 'G24'),
    ('PRESENTATION', 'G66'),
    ('INCOME STATEMENT', 'AE30'),
    ('Valuation', 'E21'),
    ('Valuation', 'E38'),
    ('Valuation', 'E43'),
    ('PRESENTATION', 'H63'),
    ('COMPANY OVERVIEW', 'H24'),
    ('PRESENTATION', 'H66'),
    ('INCOME STATEMENT', 'AF30'),
    ('Valuation', 'F21'),
    ('Valuation', 'F38'),
    ('PRESENTATION', 'I63'),
    ('COMPANY OVERVIEW', 'I24'),
    ('PRESENTATION', 'I66'),
    ('INCOME STATEMENT', 'AG30'),
    ('Valuation', 'G21'),
    ('Valuation', 'G38'),
    ('PRESENTATION', 'J63'),
    ('COMPANY OVERVIEW', 'J24'),
    ('PRESENTATION', 'J66'),
    ('INCOME STATEMENT', 'K36'),
    ('INCOME STATEMENT', 'K37'),
    ('INCOME STATEMENT', 'K38'),
    ('Valuation', 'C40'),
    ('Valuation', 'B45'),
    ('INCOME STATEMENT', 'G47'),
    ('INCOME STATEMENT', 'F48'),
    ('INCOME STATEMENT', 'X47'),
    ('INCOME STATEMENT', 'W48'),
    ('Valuation', 'B30'),
    ('CASH FOW STATEMENT', 'I23'),
    ('Valuation', 'G46'),
    ('Valuation', 'J25'),
    ('COMPANY OVERVIEW', 'E25'),
    ('COMPANY OVERVIEW', 'F25'),
    ('COMPANY OVERVIEW', 'G25'),
    ('COMPANY OVERVIEW', 'H25'),
    ('COMPANY OVERVIEW', 'I25'),
    ('COMPANY OVERVIEW', 'J25'),
    ('COMPANY OVERVIEW', 'G26'),
    ('COMPANY OVERVIEW', 'H26'),
    ('COMPANY OVERVIEW', 'I26'),
    ('COMPANY OVERVIEW', 'J26'),
    ('Valuation', 'K6'),
    ('Valuation', 'K7'),
    ('Valuation', 'H37'),
    ('Valuation', 'H45'),
    ('PRESENTATION', 'B44'),
    ('PRESENTATION', 'C44'),
    ('PRESENTATION', 'E44'),
    ('PRESENTATION', 'G44'),
    ('PRESENTATION', 'I44'),
    ('PRESENTATION', 'J44'),
    ('PRESENTATION', 'L44'),
    ('PRESENTATION', 'N44'),
    ('PRESENTATION', 'P44'),
    ('PRESENTATION', 'Q44'),
    ('PRESENTATION', 'S44'),
    ('PRESENTATION', 'U44'),
    ('PRESENTATION', 'W44'),
    ('PRESENTATION', 'X44'),
    ('PRESENTATION', 'Z44'),
    ('PRESENTATION', 'AB44'),
    ('PRESENTATION', 'AD41'),
    ('PRESENTATION', 'AE41'),
    ('PRESENTATION', 'AF41'),
    ('PRESENTATION', 'AG41'),
    ('PRESENTATION', 'H33'),
    ('PRESENTATION', 'H36'),
    ('PRESENTATION', 'O33'),
    ('PRESENTATION', 'O36'),
    ('PRESENTATION', 'D40'),
    ('PRESENTATION', 'F38'),
    ('PRESENTATION', 'M38'),
    ('PRESENTATION', 'R40'),
    ('PRESENTATION', 'T38'),
    ('PRESENTATION', 'V33'),
    ('PRESENTATION', 'V36'),
    ('PRESENTATION', 'Y40'),
    ('PRESENTATION', 'AA38'),
    ('PRESENTATION', 'AC33'),
    ('PRESENTATION', 'AC36'),
    ('PRESENTATION', 'K45'),
    ('PRESENTATION', 'K47'),
    ('Segment Revenue Model', 'Y8'),
    ('Segment Revenue Model', 'Y9'),
    ('Segment Revenue Model', 'Y10'),
    ('Segment Revenue Model', 'Y11'),
    ('Segment Revenue Model', 'Y13'),
    ('Segment Revenue Model', 'U17'),
    ('Segment Revenue Model', 'T33'),
    ('Segment Revenue Model', 'T37'),
    ('Segment Revenue Model', 'T38'),
    ('Segment Revenue Model', 'T39'),
    ('Segment Revenue Model', 'T40'),
    ('Segment Revenue Model', 'T41'),
    ('Segment Revenue Model', 'T42'),
    ('Segment Revenue Model', 'T43'),
    ('Segment Revenue Model', 'T44'),
    ('Segment Revenue Model', 'T45'),
    ('Segment Revenue Model', 'T46'),
    ('Segment Revenue Model', 'Y14'),
    ('Segment Revenue Model', 'X15'),
    ('Segment Revenue Model', 'W17'),
    ('Segment Revenue Model', 'V37'),
    ('Segment Revenue Model', 'V38'),
    ('Segment Revenue Model', 'V39'),
    ('Segment Revenue Model', 'V40'),
    ('Segment Revenue Model', 'V41'),
    ('Segment Revenue Model', 'V42'),
    ('Segment Revenue Model', 'V43'),
    ('Segment Revenue Model', 'V44'),
    ('Segment Revenue Model', 'V45'),
    ('Segment Revenue Model', 'V46'),
    ('INCOME STATEMENT', 'B36'),
    ('INCOME STATEMENT', 'B37'),
    ('INCOME STATEMENT', 'B38'),
    ('INCOME STATEMENT', 'C36'),
    ('INCOME STATEMENT', 'C37'),
    ('INCOME STATEMENT', 'C38'),
    ('INCOME STATEMENT', 'E36'),
    ('INCOME STATEMENT', 'E37'),
    ('INCOME STATEMENT', 'E38'),
    ('INCOME STATEMENT', 'G36'),
    ('INCOME STATEMENT', 'G37'),
    ('INCOME STATEMENT', 'G38'),
    ('INCOME STATEMENT', 'I36'),
    ('INCOME STATEMENT', 'I37'),
    ('INCOME STATEMENT', 'I38'),
    ('INCOME STATEMENT', 'J36'),
    ('INCOME STATEMENT', 'J37'),
    ('INCOME STATEMENT', 'J38'),
    ('INCOME STATEMENT', 'L36'),
    ('INCOME STATEMENT', 'L37'),
    ('INCOME STATEMENT', 'L38'),
    ('INCOME STATEMENT', 'N36'),
    ('INCOME STATEMENT', 'N37'),
    ('INCOME STATEMENT', 'N38'),
    ('INCOME STATEMENT', 'P36'),
    ('INCOME STATEMENT', 'P37'),
    ('INCOME STATEMENT', 'P38'),
    ('INCOME STATEMENT', 'Q36'),
    ('INCOME STATEMENT', 'Q37'),
    ('INCOME STATEMENT', 'Q38'),
    ('INCOME STATEMENT', 'S36'),
    ('INCOME STATEMENT', 'S37'),
    ('INCOME STATEMENT', 'S38'),
    ('INCOME STATEMENT', 'U36'),
    ('INCOME STATEMENT', 'U37'),
    ('INCOME STATEMENT', 'U38'),
    ('INCOME STATEMENT', 'W36'),
    ('INCOME STATEMENT', 'W37'),
    ('INCOME STATEMENT', 'W38'),
    ('INCOME STATEMENT', 'X36'),
    ('INCOME STATEMENT', 'X37'),
    ('INCOME STATEMENT', 'X38'),
    ('INCOME STATEMENT', 'Z36'),
    ('INCOME STATEMENT', 'Z37'),
    ('INCOME STATEMENT', 'Z38'),
    ('INCOME STATEMENT', 'AB36'),
    ('INCOME STATEMENT', 'AB37'),
    ('INCOME STATEMENT', 'AB38'),
    ('COMPANY OVERVIEW', 'C22'),
    ('INCOME STATEMENT', 'B22'),
    ('INCOME STATEMENT', 'C22'),
    ('INCOME STATEMENT', 'D22'),
    ('INCOME STATEMENT', 'E22'),
    ('INCOME STATEMENT', 'F22'),
    ('INCOME STATEMENT', 'G22'),
    ('INCOME STATEMENT', 'H25'),
    ('INCOME STATEMENT', 'H28'),
    ('Ratio Analysis', 'C13'),
    ('Ratio Analysis', 'C19'),
    ('COMPANY OVERVIEW', 'D22'),
    ('INCOME STATEMENT', 'I22'),
    ('INCOME STATEMENT', 'J22'),
    ('INCOME STATEMENT', 'K22'),
    ('INCOME STATEMENT', 'L22'),
    ('INCOME STATEMENT', 'M22'),
    ('INCOME STATEMENT', 'N22'),
    ('INCOME STATEMENT', 'O23'),
    ('INCOME STATEMENT', 'O25'),
    ('INCOME STATEMENT', 'O28'),
    ('Ratio Analysis', 'D13'),
    ('Ratio Analysis', 'D19'),
    ('Ratio Analysis', 'D22'),
    ('INCOME STATEMENT', 'D32'),
    ('INCOME STATEMENT', 'F30'),
    ('INCOME STATEMENT', 'M30'),
    ('INCOME STATEMENT', 'R32'),
    ('INCOME STATEMENT', 'T30'),
    ('COMPANY OVERVIEW', 'E22'),
    ('COMPANY OVERVIEW', 'E26'),
    ('INCOME STATEMENT', 'P22'),
    ('INCOME STATEMENT', 'Q22'),
    ('INCOME STATEMENT', 'R22'),
    ('INCOME STATEMENT', 'S22'),
    ('INCOME STATEMENT', 'T22'),
    ('INCOME STATEMENT', 'U22'),
    ('INCOME STATEMENT', 'V23'),
    ('INCOME STATEMENT', 'V25'),
    ('INCOME STATEMENT', 'V28'),
    ('Valuation', 'B20'),
    ('Valuation', 'B21'),
    ('Ratio Analysis', 'E13'),
    ('Ratio Analysis', 'E19'),
    ('Ratio Analysis', 'E22'),
    ('INCOME STATEMENT', 'Y32'),
    ('INCOME STATEMENT', 'AA30'),
    ('COMPANY OVERVIEW', 'F22'),
    ('COMPANY OVERVIEW', 'F26'),
    ('INCOME STATEMENT', 'W22'),
    ('INCOME STATEMENT', 'X22'),
    ('INCOME STATEMENT', 'Y22'),
    ('INCOME STATEMENT', 'Z22'),
    ('INCOME STATEMENT', 'AA22'),
    ('INCOME STATEMENT', 'AB22'),
    ('INCOME STATEMENT', 'AC23'),
    ('INCOME STATEMENT', 'AD23'),
    ('INCOME STATEMENT', 'AC25'),
    ('INCOME STATEMENT', 'AC28'),
    ('Valuation', 'C20'),
    ('Ratio Analysis', 'F13'),
    ('Ratio Analysis', 'F19'),
    ('Ratio Analysis', 'F22'),
    ('INCOME STATEMENT', 'AD32'),
    ('Ratio Analysis', 'G26'),
    ('Valuation', 'D44'),
    ('INCOME STATEMENT', 'AE32'),
    ('Ratio Analysis', 'H26'),
    ('Valuation', 'E39'),
    ('Valuation', 'E44'),
    ('INCOME STATEMENT', 'AF32'),
    ('Ratio Analysis', 'I26'),
    ('Valuation', 'F39'),
    ('INCOME STATEMENT', 'AG32'),
    ('Ratio Analysis', 'J26'),
    ('Valuation', 'G39'),
    ('INCOME STATEMENT', 'K41'),
    ('INCOME STATEMENT', 'H47'),
    ('INCOME STATEMENT', 'G48'),
    ('INCOME STATEMENT', 'Y47'),
    ('INCOME STATEMENT', 'X48'),
    ('CASH FOW STATEMENT', 'I24'),
    ('Valuation', 'H46'),
    ('Valuation', 'K25'),
    ('Valuation', 'H19'),
    ('Valuation', 'I37'),
    ('Valuation', 'I45'),
    ('PRESENTATION', 'B45'),
    ('PRESENTATION', 'C45'),
    ('PRESENTATION', 'C47'),
    ('PRESENTATION', 'E45'),
    ('PRESENTATION', 'E47'),
    ('PRESENTATION', 'G45'),
    ('PRESENTATION', 'G47'),
    ('PRESENTATION', 'I45'),
    ('PRESENTATION', 'I47'),
    ('PRESENTATION', 'J45'),
    ('PRESENTATION', 'J47'),
    ('PRESENTATION', 'L45'),
    ('PRESENTATION', 'L47'),
    ('PRESENTATION', 'N45'),
    ('PRESENTATION', 'N47'),
    ('PRESENTATION', 'P45'),
    ('PRESENTATION', 'P47'),
    ('PRESENTATION', 'Q45'),
    ('PRESENTATION', 'Q47'),
    ('PRESENTATION', 'S45'),
    ('PRESENTATION', 'S47'),
    ('PRESENTATION', 'U45'),
    ('PRESENTATION', 'U47'),
    ('PRESENTATION', 'W45'),
    ('PRESENTATION', 'W47'),
    ('PRESENTATION', 'X45'),
    ('PRESENTATION', 'X47'),
    ('PRESENTATION', 'Z45'),
    ('PRESENTATION', 'Z47'),
    ('PRESENTATION', 'AB45'),
    ('PRESENTATION', 'AB47'),
    ('PRESENTATION', 'AD44'),
    ('PRESENTATION', 'AE44'),
    ('PRESENTATION', 'AF44'),
    ('PRESENTATION', 'AG44'),
    ('PRESENTATION', 'H38'),
    ('PRESENTATION', 'O38'),
    ('PRESENTATION', 'D44'),
    ('PRESENTATION', 'F40'),
    ('PRESENTATION', 'M40'),
    ('PRESENTATION', 'R44'),
    ('PRESENTATION', 'T40'),
    ('PRESENTATION', 'V38'),
    ('PRESENTATION', 'Y44'),
    ('PRESENTATION', 'AA40'),
    ('PRESENTATION', 'AC38'),
    ('Segment Revenue Model', 'U33'),
    ('Segment Revenue Model', 'V33'),
    ('Segment Revenue Model', 'U37'),
    ('Segment Revenue Model', 'U38'),
    ('Segment Revenue Model', 'U39'),
    ('Segment Revenue Model', 'U40'),
    ('Segment Revenue Model', 'U42'),
    ('Segment Revenue Model', 'U43'),
    ('Segment Revenue Model', 'U44'),
    ('Segment Revenue Model', 'U45'),
    ('Segment Revenue Model', 'U46'),
    ('Segment Revenue Model', 'Y15'),
    ('Segment Revenue Model', 'X17'),
    ('Segment Revenue Model', 'W33'),
    ('Segment Revenue Model', 'W37'),
    ('Segment Revenue Model', 'W38'),
    ('Segment Revenue Model', 'W39'),
    ('Segment Revenue Model', 'W40'),
    ('Segment Revenue Model', 'W41'),
    ('Segment Revenue Model', 'W42'),
    ('Segment Revenue Model', 'W43'),
    ('Segment Revenue Model', 'W44'),
    ('Segment Revenue Model', 'W45'),
    ('Segment Revenue Model', 'W46'),
    ('INCOME STATEMENT', 'B41'),
    ('INCOME STATEMENT', 'B43'),
    ('INCOME STATEMENT', 'B44'),
    ('INCOME STATEMENT', 'C41'),
    ('INCOME STATEMENT', 'C43'),
    ('INCOME STATEMENT', 'C44'),
    ('INCOME STATEMENT', 'E41'),
    ('INCOME STATEMENT', 'E43'),
    ('INCOME STATEMENT', 'E44'),
    ('INCOME STATEMENT', 'G41'),
    ('INCOME STATEMENT', 'G43'),
    ('INCOME STATEMENT', 'G44'),
    ('INCOME STATEMENT', 'I41'),
    ('INCOME STATEMENT', 'I43'),
    ('INCOME STATEMENT', 'I44'),
    ('INCOME STATEMENT', 'J41'),
    ('INCOME STATEMENT', 'J43'),
    ('INCOME STATEMENT', 'J44'),
    ('INCOME STATEMENT', 'L41'),
    ('INCOME STATEMENT', 'L43'),
    ('INCOME STATEMENT', 'L44'),
    ('INCOME STATEMENT', 'N41'),
    ('INCOME STATEMENT', 'N43'),
    ('INCOME STATEMENT', 'N44'),
    ('INCOME STATEMENT', 'P41'),
    ('INCOME STATEMENT', 'P43'),
    ('INCOME STATEMENT', 'P44'),
    ('INCOME STATEMENT', 'Q41'),
    ('INCOME STATEMENT', 'Q43'),
    ('INCOME STATEMENT', 'Q44'),
    ('INCOME STATEMENT', 'S41'),
    ('INCOME STATEMENT', 'S43'),
    ('INCOME STATEMENT', 'S44'),
    ('INCOME STATEMENT', 'U41'),
    ('INCOME STATEMENT', 'U43'),
    ('INCOME STATEMENT', 'U44'),
    ('INCOME STATEMENT', 'W41'),
    ('INCOME STATEMENT', 'W43'),
    ('INCOME STATEMENT', 'W44'),
    ('INCOME STATEMENT', 'X41'),
    ('INCOME STATEMENT', 'X43'),
    ('INCOME STATEMENT', 'X44'),
    ('INCOME STATEMENT', 'Z41'),
    ('INCOME STATEMENT', 'Z43'),
    ('INCOME STATEMENT', 'Z44'),
    ('INCOME STATEMENT', 'AB41'),
    ('INCOME STATEMENT', 'AB43'),
    ('INCOME STATEMENT', 'AB44'),
    ('INCOME STATEMENT', 'H22'),
    ('INCOME STATEMENT', 'H30'),
    ('PRESENTATION', 'C63'),
    ('COMPANY OVERVIEW', 'C24'),
    ('PRESENTATION', 'C66'),
    ('INCOME STATEMENT', 'O22'),
    ('INCOME STATEMENT', 'O30'),
    ('PRESENTATION', 'D63'),
    ('COMPANY OVERVIEW', 'D24'),
    ('PRESENTATION', 'D66'),
    ('INCOME STATEMENT', 'D36'),
    ('INCOME STATEMENT', 'D37'),
    ('INCOME STATEMENT', 'D38'),
    ('INCOME STATEMENT', 'F32'),
    ('INCOME STATEMENT', 'M32'),
    ('INCOME STATEMENT', 'R36'),
    ('INCOME STATEMENT', 'R37'),
    ('INCOME STATEMENT', 'R38'),
    ('INCOME STATEMENT', 'T32'),
    ('INCOME STATEMENT', 'V22'),
    ('INCOME STATEMENT', 'V30'),
    ('Valuation', 'B43'),
    ('Valuation', 'B23'),
    ('Valuation', 'B44'),
    ('Valuation', 'B48'),
    ('PRESENTATION', 'E63'),
    ('COMPANY OVERVIEW', 'E24'),
    ('PRESENTATION', 'E66'),
    ('INCOME STATEMENT', 'Y36'),
    ('INCOME STATEMENT', 'Y37'),
    ('INCOME STATEMENT', 'Y38'),
    ('INCOME STATEMENT', 'AA32'),
    ('INCOME STATEMENT', 'AC22'),
    ('INCOME STATEMENT', 'AC30'),
    ('Valuation', 'C21'),
    ('Valuation', 'C38'),
    ('Valuation', 'D38'),
    ('Valuation', 'C43'),
    ('PRESENTATION', 'F63'),
    ('COMPANY OVERVIEW', 'F24'),
    ('PRESENTATION', 'F66'),
    ('INCOME STATEMENT', 'AD33'),
    ('Ratio Analysis', 'G12'),
    ('Ratio Analysis', 'G18'),
    ('INCOME STATEMENT', 'AE33'),
    ('CASH FOW STATEMENT', 'F8'),
    ('Ratio Analysis', 'H12'),
    ('Ratio Analysis', 'H18'),
    ('Valuation', 'F44'),
    ('INCOME STATEMENT', 'AF33'),
    ('CASH FOW STATEMENT', 'G8'),
    ('Ratio Analysis', 'I12'),
    ('Ratio Analysis', 'I18'),
    ('INCOME STATEMENT', 'AG33'),
    ('CASH FOW STATEMENT', 'H8'),
    ('Ratio Analysis', 'J12'),
    ('Ratio Analysis', 'J18'),
    ('INCOME STATEMENT', 'K45'),
    ('INCOME STATEMENT', 'I47'),
    ('INCOME STATEMENT', 'H48'),
    ('INCOME STATEMENT', 'Z47'),
    ('INCOME STATEMENT', 'Y48'),
    ('Valuation', 'I46'),
    ('Valuation', 'H24'),
    ('Valuation', 'H27'),
    ('Valuation', 'I19'),
    ('Valuation', 'J45'),
    ('PRESENTATION', 'B47'),
    ('PRESENTATION', 'AD45'),
    ('PRESENTATION', 'AD47'),
    ('PRESENTATION', 'AE45'),
    ('PRESENTATION', 'AE47'),
    ('PRESENTATION', 'AF45'),
    ('PRESENTATION', 'AF47'),
    ('PRESENTATION', 'AG45'),
    ('PRESENTATION', 'AG47'),
    ('PRESENTATION', 'H40'),
    ('PRESENTATION', 'O40'),
    ('PRESENTATION', 'D45'),
    ('PRESENTATION', 'D47'),
    ('PRESENTATION', 'F44'),
    ('PRESENTATION', 'M44'),
    ('PRESENTATION', 'R45'),
    ('PRESENTATION', 'R47'),
    ('PRESENTATION', 'T44'),
    ('PRESENTATION', 'V40'),
    ('PRESENTATION', 'Y45'),
    ('PRESENTATION', 'Y47'),
    ('PRESENTATION', 'AA44'),
    ('PRESENTATION', 'AC40'),
    ('Segment Revenue Model', 'Y17'),
    ('Segment Revenue Model', 'X33'),
    ('Segment Revenue Model', 'X37'),
    ('Segment Revenue Model', 'X38'),
    ('Segment Revenue Model', 'X39'),
    ('Segment Revenue Model', 'X40'),
    ('Segment Revenue Model', 'X41'),
    ('Segment Revenue Model', 'X42'),
    ('Segment Revenue Model', 'X43'),
    ('Segment Revenue Model', 'X44'),
    ('Segment Revenue Model', 'X45'),
    ('Segment Revenue Model', 'X46'),
    ('INCOME STATEMENT', 'B45'),
    ('INCOME STATEMENT', 'B49'),
    ('INCOME STATEMENT', 'C45'),
    ('INCOME STATEMENT', 'C49'),
    ('INCOME STATEMENT', 'E45'),
    ('INCOME STATEMENT', 'E49'),
    ('INCOME STATEMENT', 'G45'),
    ('INCOME STATEMENT', 'G49'),
    ('INCOME STATEMENT', 'I45'),
    ('INCOME STATEMENT', 'J45'),
    ('INCOME STATEMENT', 'L45'),
    ('INCOME STATEMENT', 'N45'),
    ('INCOME STATEMENT', 'P45'),
    ('INCOME STATEMENT', 'Q45'),
    ('INCOME STATEMENT', 'S45'),
    ('INCOME STATEMENT', 'S49'),
    ('INCOME STATEMENT', 'U45'),
    ('INCOME STATEMENT', 'U49'),
    ('INCOME STATEMENT', 'W45'),
    ('INCOME STATEMENT', 'W49'),
    ('INCOME STATEMENT', 'X45'),
    ('INCOME STATEMENT', 'X49'),
    ('INCOME STATEMENT', 'Z45'),
    ('INCOME STATEMENT', 'AB45'),
    ('INCOME STATEMENT', 'H32'),
    ('INCOME STATEMENT', 'O32'),
    ('Ratio Analysis', 'D26'),
    ('INCOME STATEMENT', 'D41'),
    ('INCOME STATEMENT', 'D43'),
    ('INCOME STATEMENT', 'K43'),
    ('INCOME STATEMENT', 'F36'),
    ('INCOME STATEMENT', 'F37'),
    ('INCOME STATEMENT', 'F38'),
    ('INCOME STATEMENT', 'M36'),
    ('INCOME STATEMENT', 'M37'),
    ('INCOME STATEMENT', 'M38'),
    ('INCOME STATEMENT', 'R41'),
    ('INCOME STATEMENT', 'R43'),
    ('INCOME STATEMENT', 'T36'),
    ('INCOME STATEMENT', 'T37'),
    ('INCOME STATEMENT', 'T38'),
    ('INCOME STATEMENT', 'V32'),
    ('Ratio Analysis', 'E26'),
    ('Valuation', 'B26'),
    ('INCOME STATEMENT', 'Y41'),
    ('INCOME STATEMENT', 'Y43'),
    ('INCOME STATEMENT', 'AA36'),
    ('INCOME STATEMENT', 'AA37'),
    ('INCOME STATEMENT', 'AA38'),
    ('INCOME STATEMENT', 'AC32'),
    ('Ratio Analysis', 'F26'),
    ('Valuation', 'C23'),
    ('Valuation', 'C39'),
    ('Valuation', 'D39'),
    ('Valuation', 'C44'),
    ('Valuation', 'C48'),
    ('INCOME STATEMENT', 'AD36'),
    ('INCOME STATEMENT', 'AD37'),
    ('INCOME STATEMENT', 'AD38'),
    ('Valuation', 'D22'),
    ('INCOME STATEMENT', 'AE36'),
    ('INCOME STATEMENT', 'AE37'),
    ('INCOME STATEMENT', 'AE38'),
    ('CASH FOW STATEMENT', 'F9'),
    ('Valuation', 'E22'),
    ('Valuation', 'G44'),
    ('INCOME STATEMENT', 'AF36'),
    ('INCOME STATEMENT', 'AF37'),
    ('INCOME STATEMENT', 'AF38'),
    ('CASH FOW STATEMENT', 'G9'),
    ('Valuation', 'F22'),
    ('INCOME STATEMENT', 'AG36'),
    ('INCOME STATEMENT', 'AG37'),
    ('INCOME STATEMENT', 'AG38'),
    ('CASH FOW STATEMENT', 'H9'),
    ('Valuation', 'G22'),
    ('INCOME STATEMENT', 'J47'),
    ('INCOME STATEMENT', 'I48'),
    ('INCOME STATEMENT', 'AA47'),
    ('INCOME STATEMENT', 'Z48'),
    ('Valuation', 'J46'),
    ('Valuation', 'H40'),
    ('Valuation', 'J19'),
    ('Valuation', 'I24'),
    ('Valuation', 'I27'),
    ('Valuation', 'K45'),
    ('PRESENTATION', 'H44'),
    ('PRESENTATION', 'O44'),
    ('PRESENTATION', 'F45'),
    ('PRESENTATION', 'F47'),
    ('PRESENTATION', 'M45'),
    ('PRESENTATION', 'M47'),
    ('PRESENTATION', 'T45'),
    ('PRESENTATION', 'T47'),
    ('PRESENTATION', 'V44'),
    ('PRESENTATION', 'AA45'),
    ('PRESENTATION', 'AA47'),
    ('PRESENTATION', 'AC44'),
    ('Segment Revenue Model', 'Y33'),
    ('Segment Revenue Model', 'Y37'),
    ('Segment Revenue Model', 'Y38'),
    ('Segment Revenue Model', 'Y39'),
    ('Segment Revenue Model', 'Y40'),
    ('Segment Revenue Model', 'Y41'),
    ('Segment Revenue Model', 'Y42'),
    ('Segment Revenue Model', 'Y43'),
    ('Segment Revenue Model', 'Y44'),
    ('Segment Revenue Model', 'Y45'),
    ('Segment Revenue Model', 'Y46'),
    ('INCOME STATEMENT', 'H36'),
    ('INCOME STATEMENT', 'H37'),
    ('INCOME STATEMENT', 'H38'),
    ('CASH FOW STATEMENT', 'B8'),
    ('Ratio Analysis', 'C12'),
    ('Ratio Analysis', 'C18'),
    ('INCOME STATEMENT', 'O36'),
    ('INCOME STATEMENT', 'O37'),
    ('INCOME STATEMENT', 'O38'),
    ('CASH FOW STATEMENT', 'C8'),
    ('Ratio Analysis', 'D12'),
    ('Ratio Analysis', 'D18'),
    ('INCOME STATEMENT', 'D45'),
    ('INCOME STATEMENT', 'D49'),
    ('INCOME STATEMENT', 'F41'),
    ('INCOME STATEMENT', 'F43'),
    ('INCOME STATEMENT', 'M41'),
    ('INCOME STATEMENT', 'M43'),
    ('INCOME STATEMENT', 'R45'),
    ('INCOME STATEMENT', 'T41'),
    ('INCOME STATEMENT', 'T43'),
    ('INCOME STATEMENT', 'V36'),
    ('INCOME STATEMENT', 'V37'),
    ('INCOME STATEMENT', 'V38'),
    ('CASH FOW STATEMENT', 'D8'),
    ('Ratio Analysis', 'E12'),
    ('Ratio Analysis', 'E18'),
    ('Valuation', 'B28'),
    ('INCOME STATEMENT', 'Y45'),
    ('INCOME STATEMENT', 'Y49'),
    ('INCOME STATEMENT', 'AA41'),
    ('INCOME STATEMENT', 'AA43'),
    ('INCOME STATEMENT', 'AC36'),
    ('INCOME STATEMENT', 'AC37'),
    ('INCOME STATEMENT', 'AC38'),
    ('CASH FOW STATEMENT', 'E8'),
    ('Ratio Analysis', 'F12'),
    ('Ratio Analysis', 'F18'),
    ('Valuation', 'C26'),
    ('Ratio Analysis', 'G25'),
    ('Ratio Analysis', 'G16'),
    ('Ratio Analysis', 'G11'),
    ('Valuation', 'D23'),
    ('Valuation', 'D48'),
    ('Ratio Analysis', 'H25'),
    ('Ratio Analysis', 'H16'),
    ('Ratio Analysis', 'H11'),
    ('CASH FOW STATEMENT', 'F10'),
    ('Valuation', 'E23'),
    ('Valuation', 'E48'),
    ('Valuation', 'H44'),
    ('Ratio Analysis', 'I25'),
    ('Ratio Analysis', 'I16'),
    ('Ratio Analysis', 'I11'),
    ('CASH FOW STATEMENT', 'G10'),
    ('Valuation', 'F23'),
    ('Ratio Analysis', 'J25'),
    ('Ratio Analysis', 'J16'),
    ('Ratio Analysis', 'J11'),
    ('CASH FOW STATEMENT', 'H10'),
    ('Valuation', 'G23'),
    ('INCOME STATEMENT', 'K47'),
    ('INCOME STATEMENT', 'J48'),
    ('INCOME STATEMENT', 'I49'),
    ('INCOME STATEMENT', 'AB47'),
    ('INCOME STATEMENT', 'AA48'),
    ('INCOME STATEMENT', 'Z49'),
    ('Valuation', 'K46'),
    ('Valuation', 'K19'),
    ('Valuation', 'J24'),
    ('Valuation', 'J27'),
    ('Valuation', 'I40'),
    ('PRESENTATION', 'H45'),
    ('PRESENTATION', 'H47'),
    ('PRESENTATION', 'O45'),
    ('PRESENTATION', 'O47'),
    ('PRESENTATION', 'V45'),
    ('PRESENTATION', 'V47'),
    ('PRESENTATION', 'AC45'),
    ('PRESENTATION', 'AC47'),
    ('Ratio Analysis', 'C25'),
    ('Ratio Analysis', 'C16'),
    ('INCOME STATEMENT', 'H41'),
    ('INCOME STATEMENT', 'B42'),
    ('INCOME STATEMENT', 'C42'),
    ('INCOME STATEMENT', 'D42'),
    ('INCOME STATEMENT', 'E42'),
    ('INCOME STATEMENT', 'F42'),
    ('INCOME STATEMENT', 'G42'),
    ('INCOME STATEMENT', 'H43'),
    ('Ratio Analysis', 'C11'),
    ('CASH FOW STATEMENT', 'B10'),
    ('Ratio Analysis', 'D25'),
    ('Ratio Analysis', 'D16'),
    ('INCOME STATEMENT', 'O41'),
    ('INCOME STATEMENT', 'O43'),
    ('Ratio Analysis', 'D11'),
    ('CASH FOW STATEMENT', 'C10'),
    ('INCOME STATEMENT', 'F45'),
    ('INCOME STATEMENT', 'F49'),
    ('INCOME STATEMENT', 'M45'),
    ('INCOME STATEMENT', 'T45'),
    ('INCOME STATEMENT', 'T49'),
    ('Ratio Analysis', 'E25'),
    ('Valuation', 'K10'),
    ('Ratio Analysis', 'E16'),
    ('INCOME STATEMENT', 'V41'),
    ('INCOME STATEMENT', 'P42'),
    ('INCOME STATEMENT', 'Q42'),
    ('INCOME STATEMENT', 'R42'),
    ('INCOME STATEMENT', 'S42'),
    ('INCOME STATEMENT', 'T42'),
    ('INCOME STATEMENT', 'U42'),
    ('INCOME STATEMENT', 'V43'),
    ('Ratio Analysis', 'E11'),
    ('CASH FOW STATEMENT', 'D10'),
    ('INCOME STATEMENT', 'AC41'),
    ('INCOME STATEMENT', 'AA45'),
    ('Ratio Analysis', 'F25'),
    ('Ratio Analysis', 'F16'),
    ('INCOME STATEMENT', 'AC43'),
    ('Ratio Analysis', 'F11'),
    ('CASH FOW STATEMENT', 'E10'),
    ('Valuation', 'C28'),
    ('Valuation', 'D26'),
    ('CASH FOW STATEMENT', 'F12'),
    ('Valuation', 'E26'),
    ('Valuation', 'F48'),
    ('Valuation', 'H21'),
    ('Valuation', 'I44'),
    ('CASH FOW STATEMENT', 'G12'),
    ('Valuation', 'F26'),
    ('CASH FOW STATEMENT', 'H12'),
    ('Valuation', 'G26'),
    ('INCOME STATEMENT', 'L47'),
    ('INCOME STATEMENT', 'K48'),
    ('INCOME STATEMENT', 'J49'),
    ('INCOME STATEMENT', 'AC47'),
    ('INCOME STATEMENT', 'AB48'),
    ('INCOME STATEMENT', 'AA49'),
    ('Valuation', 'K24'),
    ('Valuation', 'K27'),
    ('Valuation', 'J40'),
    ('INCOME STATEMENT', 'H45'),
    ('INCOME STATEMENT', 'H49'),
    ('Ratio Analysis', 'C14'),
    ('INCOME STATEMENT', 'H42'),
    ('CASH FOW STATEMENT', 'B12'),
    ('COMPANY OVERVIEW', 'C23'),
    ('INCOME STATEMENT', 'I42'),
    ('INCOME STATEMENT', 'J42'),
    ('INCOME STATEMENT', 'K42'),
    ('INCOME STATEMENT', 'L42'),
    ('INCOME STATEMENT', 'M42'),
    ('INCOME STATEMENT', 'N42'),
    ('INCOME STATEMENT', 'O45'),
    ('Ratio Analysis', 'D14'),
    ('CASH FOW STATEMENT', 'C12'),
    ('Valuation', 'K13'),
    ('COMPANY OVERVIEW', 'D23'),
    ('COMPANY OVERVIEW', 'E23'),
    ('INCOME STATEMENT', 'V45'),
    ('INCOME STATEMENT', 'V49'),
    ('Ratio Analysis', 'E14'),
    ('INCOME STATEMENT', 'V42'),
    ('CASH FOW STATEMENT', 'D12'),
    ('COMPANY OVERVIEW', 'F23'),
    ('INCOME STATEMENT', 'W42'),
    ('INCOME STATEMENT', 'X42'),
    ('INCOME STATEMENT', 'Y42'),
    ('INCOME STATEMENT', 'Z42'),
    ('INCOME STATEMENT', 'AA42'),
    ('INCOME STATEMENT', 'AB42'),
    ('INCOME STATEMENT', 'AD43'),
    ('INCOME STATEMENT', 'AC45'),
    ('Ratio Analysis', 'F14'),
    ('CASH FOW STATEMENT', 'E12'),
    ('Valuation', 'D28'),
    ('CASH FOW STATEMENT', 'F14'),
    ('Valuation', 'E28'),
    ('Valuation', 'G48'),
    ('Valuation', 'H20'),
    ('Valuation', 'H39'),
    ('Valuation', 'I21'),
    ('Valuation', 'J44'),
    ('CASH FOW STATEMENT', 'G14'),
    ('Valuation', 'F28'),
    ('CASH FOW STATEMENT', 'H14'),
    ('Valuation', 'G28'),
    ('INCOME STATEMENT', 'M47'),
    ('INCOME STATEMENT', 'L48'),
    ('INCOME STATEMENT', 'K49'),
    ('INCOME STATEMENT', 'AD47'),
    ('INCOME STATEMENT', 'AC48'),
    ('INCOME STATEMENT', 'AB49'),
    ('Valuation', 'K40'),
    ('COMPANY OVERVIEW', 'C21'),
    ('Ratio Analysis', 'C27'),
    ('PRESENTATION', 'C65'),
    ('CASH FOW STATEMENT', 'B14'),
    ('INCOME STATEMENT', 'O42'),
    ('PRESENTATION', 'D65'),
    ('CASH FOW STATEMENT', 'C14'),
    ('Valuation', 'C33'),
    ('Valuation', 'D33'),
    ('Valuation', 'E33'),
    ('Valuation', 'F33'),
    ('Valuation', 'G33'),
    ('Valuation', 'H33'),
    ('Valuation', 'I33'),
    ('Valuation', 'J33'),
    ('Valuation', 'B53'),
    ('COMPANY OVERVIEW', 'E21'),
    ('COMPANY OVERVIEW', 'E27'),
    ('Ratio Analysis', 'E27'),
    ('PRESENTATION', 'E65'),
    ('CASH FOW STATEMENT', 'D14'),
    ('INCOME STATEMENT', 'AC42'),
    ('PRESENTATION', 'F65'),
    ('CASH FOW STATEMENT', 'E14'),
    ('CASH FOW STATEMENT', 'F23'),
    ('Valuation', 'H48'),
    ('Valuation', 'H38'),
    ('Valuation', 'I20'),
    ('Valuation', 'I39'),
    ('Valuation', 'J21'),
    ('Valuation', 'K44'),
    ('CASH FOW STATEMENT', 'G23'),
    ('CASH FOW STATEMENT', 'H23'),
    ('INCOME STATEMENT', 'N47'),
    ('INCOME STATEMENT', 'M48'),
    ('INCOME STATEMENT', 'L49'),
    ('INCOME STATEMENT', 'AE47'),
    ('INCOME STATEMENT', 'AD48'),
    ('INCOME STATEMENT', 'AC49'),
    ('Valuation', 'C30'),
    ('CASH FOW STATEMENT', 'B23'),
    ('CASH FOW STATEMENT', 'C23'),
    ('Valuation', 'C34'),
    ('Valuation', 'D34'),
    ('Valuation', 'E34'),
    ('Valuation', 'F34'),
    ('Valuation', 'G34'),
    ('Valuation', 'A67'),
    ('Valuation', 'A77'),
    ('CASH FOW STATEMENT', 'D23'),
    ('CASH FOW STATEMENT', 'E23'),
    ('CASH FOW STATEMENT', 'F24'),
    ('Valuation', 'H22'),
    ('Valuation', 'I48'),
    ('Valuation', 'I38'),
    ('Valuation', 'J20'),
    ('Valuation', 'J39'),
    ('Valuation', 'K21'),
    ('CASH FOW STATEMENT', 'G24'),
    ('CASH FOW STATEMENT', 'H24'),
    ('INCOME STATEMENT', 'O47'),
    ('INCOME STATEMENT', 'P47'),
    ('INCOME STATEMENT', 'N48'),
    ('INCOME STATEMENT', 'M49'),
    ('INCOME STATEMENT', 'AF47'),
    ('INCOME STATEMENT', 'AE48'),
    ('INCOME STATEMENT', 'AD49'),
    ('Valuation', 'D30'),
    ('COMPANY OVERVIEW', 'F21'),
    ('COMPANY OVERVIEW', 'F27'),
    ('Ratio Analysis', 'F27'),
    ('CASH FOW STATEMENT', 'E24'),
    ('Valuation', 'H23'),
    ('Valuation', 'I22'),
    ('Valuation', 'J48'),
    ('Valuation', 'J38'),
    ('Valuation', 'K20'),
    ('Valuation', 'K39'),
    ('Valuation', 'D75'),
    ('Valuation', 'E75'),
    ('Valuation', 'F75'),
    ('Valuation', 'D76'),
    ('Valuation', 'E76'),
    ('Valuation', 'F76'),
    ('Valuation', 'D77'),
    ('Valuation', 'E77'),
    ('Valuation', 'F77'),
    ('Valuation', 'D78'),
    ('Valuation', 'E78'),
    ('Valuation', 'F78'),
    ('Valuation', 'D79'),
    ('Valuation', 'E79'),
    ('Valuation', 'F79'),
    ('INCOME STATEMENT', 'O48'),
    ('INCOME STATEMENT', 'Q47'),
    ('INCOME STATEMENT', 'P48'),
    ('INCOME STATEMENT', 'N49'),
    ('INCOME STATEMENT', 'AG47'),
    ('INCOME STATEMENT', 'AF48'),
    ('INCOME STATEMENT', 'AE49'),
    ('Valuation', 'E30'),
    ('COMPANY OVERVIEW', 'G21'),
    ('COMPANY OVERVIEW', 'G27'),
    ('Ratio Analysis', 'G27'),
    ('Valuation', 'H26'),
    ('Valuation', 'I23'),
    ('Valuation', 'J22'),
    ('Valuation', 'K48'),
    ('Valuation', 'K38'),
    ('INCOME STATEMENT', 'O49'),
    ('INCOME STATEMENT', 'R47'),
    ('INCOME STATEMENT', 'Q48'),
    ('INCOME STATEMENT', 'P49'),
    ('INCOME STATEMENT', 'AG48'),
    ('INCOME STATEMENT', 'AF49'),
    ('Valuation', 'F30'),
    ('COMPANY OVERVIEW', 'H21'),
    ('COMPANY OVERVIEW', 'H27'),
    ('Ratio Analysis', 'H27'),
    ('Valuation', 'H28'),
    ('Valuation', 'I26'),
    ('Valuation', 'J23'),
    ('Valuation', 'K22'),
    ('COMPANY OVERVIEW', 'D21'),
    ('Ratio Analysis', 'D27'),
    ('INCOME STATEMENT', 'R48'),
    ('INCOME STATEMENT', 'Q49'),
    ('INCOME STATEMENT', 'AG49'),
    ('Valuation', 'G30'),
    ('COMPANY OVERVIEW', 'I21'),
    ('COMPANY OVERVIEW', 'I27'),
    ('Ratio Analysis', 'I27'),
    ('Valuation', 'H34'),
    ('Valuation', 'I28'),
    ('Valuation', 'J26'),
    ('Valuation', 'K23'),
    ('INCOME STATEMENT', 'R49'),
    ('COMPANY OVERVIEW', 'J21'),
    ('COMPANY OVERVIEW', 'J27'),
    ('Ratio Analysis', 'J27'),
    ('Valuation', 'H30'),
    ('Valuation', 'I34'),
    ('Valuation', 'J28'),
    ('Valuation', 'K26'),
    ('Valuation', 'I30'),
    ('Valuation', 'J34'),
    ('Valuation', 'B65'),
    ('Valuation', 'B66'),
    ('Valuation', 'B67'),
    ('Valuation', 'B68'),
    ('Valuation', 'B69'),
    ('Valuation', 'B75'),
    ('Valuation', 'B76'),
    ('Valuation', 'B77'),
    ('Valuation', 'B78'),
    ('Valuation', 'B79'),
    ('Valuation', 'K28'),
    ('Valuation', 'J30'),
    ('Valuation', 'B52'),
    ('Valuation', 'H65'),
    ('Valuation', 'I65'),
    ('Valuation', 'J65'),
    ('Valuation', 'H66'),
    ('Valuation', 'I66'),
    ('Valuation', 'J66'),
    ('Valuation', 'H67'),
    ('Valuation', 'I67'),
    ('Valuation', 'J67'),
    ('Valuation', 'H68'),
    ('Valuation', 'I68'),
    ('Valuation', 'J68'),
    ('Valuation', 'H69'),
    ('Valuation', 'I69'),
    ('Valuation', 'J69'),
    ('Valuation', 'H75'),
    ('Valuation', 'I75'),
    ('Valuation', 'J75'),
    ('Valuation', 'H76'),
    ('Valuation', 'I76'),
    ('Valuation', 'J76'),
    ('Valuation', 'H77'),
    ('Valuation', 'I77'),
    ('Valuation', 'J77'),
    ('Valuation', 'H78'),
    ('Valuation', 'I78'),
    ('Valuation', 'J78'),
    ('Valuation', 'H79'),
    ('Valuation', 'I79'),
    ('Valuation', 'J79'),
    ('Valuation', 'B55'),
    ('Valuation', 'K30'),
    ('Valuation', 'G52'),
    ('Valuation', 'B56'),
    ('Valuation', 'C64'),
    ('Valuation', 'G56'),
    ('Valuation', 'K52'),
    ('Valuation', 'K55'),
]

def run(mapping_report_path: Path = DEFAULT_MAPPING_REPORT, inputs_path: Path = DEFAULT_INPUT, output_path: Path = DEFAULT_OUTPUT) -> Path:
    return execute_unstructured_python_engine(
        mapping_report_path=mapping_report_path,
        unstructured_input_path=inputs_path,
        output_path=output_path,
        formula_order=FORMULA_ORDER,
        formula_funcs=FORMULA_FUNCS,
    )

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Run generated Python formula engine on unstructured inputs"
    )
    parser.add_argument('--mapping', type=Path, default=DEFAULT_MAPPING_REPORT)
    parser.add_argument('--inputs', type=Path, default=DEFAULT_INPUT)
    parser.add_argument('--output', type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()
    run(mapping_report_path=args.mapping, inputs_path=args.inputs, output_path=args.output)

if __name__ == '__main__':
    main()
