import sys
import json

from service.xl_anlyzer import XlAnalyzer
from util.com_utils import ComUtils

args = sys.argv
if len(args) < 2 or args[1] is None or args[2] is None:
    exit(1)

file_name = args[1]
sheet_name = args[2]

analyzer = XlAnalyzer(file_name, sheet_name)

output_file_name = analyzer.get_output_file_name("json")
table_name = analyzer.get_table_name()
f = open(output_file_name, mode='w', encoding='utf-8')

records = []
for row in analyzer.get_record_range():
    dict_record = analyzer.create_dict(row)
    print(dict_record)
    records.append(dict_record)
output = {}
output[table_name] = records
f.write(json.dumps(output, indent=2, ensure_ascii=False))

f.close

