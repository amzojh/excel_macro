import xlwings as xw
import pandas as pd
import numpy as np


class excelUtil():
    
    def __init__(self):
        pass

    
    def subtotalModule(self, x):
        pass


    def _get_column_distance_from_starting_points(self, x):
        length = len(x)
        accum_sum = 0

        print(x)

        for i in range(length):
            lower_string = x[i].lower()
            number = ord(lower_string) - 96
            number = number * (26 ** (length - i - 1))
            accum_sum =+ number 

        return accum_sum

    def get_distance_between_two_ranges(self, x, y):

        # while x.offset(0, offset)
        print(x, y)
        x_coordinates = self._parsing_addreese(x)
        y_coordinates = self._parsing_addreese(y)

        column_x = self._get_column_distance_from_starting_points(x_coordinates[0])
        column_y = self._get_column_distance_from_starting_points(y_coordinates[0])


        row_x = int(x_coordinates[1])
        row_y = int(y_coordinates[1])

        return [row_y - row_x, column_y - column_x]

    def _parsing_addreese(self, x):
        address = x.get_address()
        row_columns_coordinates = address.split('$')
        print("row_columns_coordinates ", row_columns_coordinates)
        print(address)
        return row_columns_coordinates[1:]
