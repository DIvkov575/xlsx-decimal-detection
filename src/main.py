import os
from decimal_processing_xlsx_vesmar import process, process_all

# for f ile in os.listdir("input"):
#     print(file)
#     process(file, "input/", "output/")

process_all("input/", "output/", False)
