import plan_eater.scripts as scripts
import argparse

parser = argparse.ArgumentParser()
parser.add_argument('-o', '--operation', choices=['new', 'update', 'amnet', 'report'], required=True)
parser.add_argument('-f', '--file', required=False)
parser.add_argument('-w', '--week', choices=['previous', range(1,52)], required=False)
res = parser.parse_args()
if res.operation == 'new':
    scripts.create_new_json(res.file)
if res.operation == 'update':
    scripts.update_json(res.file)
if res.operation == 'amnet':
    scripts.parse_amnet()
if res.operation == 'report':
    scripts.create_report(res.week)



