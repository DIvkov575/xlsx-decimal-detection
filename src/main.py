"""
Detect Missing Decimal Points
"""
import os
from decimal_processing_xlsx_vesmar import process, process_all


def correct_decimals():
    files = os.listdir("input")
    # print(files, '\n')
    # print(os.getcwd())

    # input_path = os.path.join(os.getcwd(), "input/")
    # output_path = os.path.join(os.getcwd(), "output/")

    # print(output_path)

    # for file in files:
    #     # print(file)
    #
    #     print(input_path + " ---- " + output_path)
    #
    #     process(file, input_path, output_path)
    process_all("input/", "output/")

correct_decimals()
