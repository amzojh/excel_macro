import os
import sys

import xlwings as xw
import pandas as pd

from Util.excelMacro import excelUtil

excel_util_instance = excelUtil()
xw_func_list = ["auto_subtotal"]

@xw.sub
@xw.arg('xl_app', vba='Application')
def make_shortcut(xl_app):
    xl_app.OnKey("+^c", "value_to_formula")


@xw.sub
def value_to_formula():
    wb = xw.Book.caller()
    active_ranges = wb.app.selection.options(ndim=1)    

    # Formula와 value가 다른 경우와 value가 너무 적은경우, 0같은 경우에는 X
    for cell in active_ranges:
        print(f'address : {cell.get_address()}')
        for formula_name in xw_func_list:
            if formula_name in cell.api.Formula:
                if cell.value != "" and not cell.api.Formula == cell.value: 
                    cell.api.Formula = cell.value
                elif not cell.api.Formula==cell.value:
                    cell.api.Formula = ""
                    cell.value = ""

@xw.func
@xw.arg('x', xw.Range)
@xw.arg('range_list', xw.Range, ndim=1)
def auto_subtotal(x, range_list):

    wb = xw.Book.caller()
    active_range = wb.app.selection.options(ndim=1)[0]

    (row_distance, column_distance) = excel_util_instance.get_distance_between_two_ranges(active_range, range_list[0])
    
    print(row_distance, column_distance)

    curr_indent_level = x.api.IndentLevel
    curr_address = x.api.address
    sliced_range = range_slicer_by_address(curr_address, range_list)

    # print(f"1 : sliced_range : {sliced_range}")
    if sliced_range == None:
        return ""

    subtotal_address = get_subtotal_ranges_address(curr_indent_level, sliced_range, column_distance)

    if subtotal_address == None:
        return ""
    else:
        return f"=subtotal(9,{subtotal_address})"


def get_subtotal_ranges_address(curr_indent ,sliced_ranges, column_offset):
    base_range = sliced_ranges[0]
    last_range = None
    for row in sliced_ranges:
        row_indent_level = row.api.IndentLevel
        if curr_indent < row_indent_level:
            last_range = row
        elif curr_indent == row_indent_level:
            break

    # print(f"base range : {base_range} \n last_range : {last_range}")

    if last_range == None:
        return None
    else:
        return xw.Range(base_range, last_range).offset(0,-column_offset).get_address(False, False)
    

def range_slicer_by_address(base_address, range_list):
    range_count = range_list.count
    index = 0
    isFound = False
    while not isFound:
        # print(index)
        if (index + 1) == range_count:
            isFound = True

        if base_address == range_list[index].address:
            # print(f"base_address {base_address}")
            index = index + 1
            length_ranges = len(range_list[index:])
            if length_ranges == 0:
                return None
            else:
                return range_list[index:]
    
        index = index + 1
    return None