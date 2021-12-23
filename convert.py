import argparse
from converter import Converter

parser = argparse.ArgumentParser(description="convert excel to json format")
parser.add_argument('--source_path', type=str, required=True, help="path to excel file")
parser.add_argument('--sheet_name', type=str, required=True, help="name of the sheet")
parser.add_argument('--target_path', type=str, required=False, help="path+name of the target json file")
parser.add_argument('--orient', type=str, required=False, default='records', help="orient type: index, records or columns")

args = parser.parse_args()

json_data = Converter.excel2json(args.source_path, args.sheet_name,
                                 output_filename=args.target_path,
                                 orient=args.orient)

print(json_data)