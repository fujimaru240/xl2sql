import sys

from service.xl_anlyzer import XlAnalyzer
from util.com_utils import ComUtils

def create_bulk(row, target_col, bulk, file):
    for bulk_elem in bulk:
        query = analyzer.create_insert_sql(row, target_col, bulk_elem)
        print(query)
        file.write("{0}\n".format(query))


args = sys.argv
if len(args) < 2 or args[1] is None or args[2] is None:
    exit(1)

file_name = args[1]
sheet_name = args[2]
analyzer = XlAnalyzer(file_name, sheet_name)

analyzer_bulk = analyzer.read_bulk_setting(file_name)
bulk_data = analyzer_bulk.get_bulk_data()

output_file_name = analyzer.get_output_file_name("sql")
f = open(output_file_name, mode='w')

for row in analyzer.get_record_range():
    if bulk_data is None:
        query = analyzer.create_insert_sql(row)
        print(query)
        f.write("{0}\n".format(query))
        continue

    for k, v in bulk_data.items():
        create_bulk(row, k, v, f)

f.close

