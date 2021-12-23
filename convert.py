import argparse
from converter import Converter

json_data = Converter.excel2json('data.xlsx', 'Marks',
                                 output_filename="new_output.jxon",
                                 orient='columns')

print(json_data)