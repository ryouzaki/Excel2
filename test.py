import plan_eater.scripts as scripts
import argparse

parser = argparse.ArgumentParser()
parser.add_argument('-o', '--operation', choices=['new', 'update'], required=True)
parser.add_argument('-f', '--file', required=True)
res = parser.parse_args()
if res.operation == 'new':
    scripts.create_new_json(res.file)
if res.operation == 'update':
    scripts.update_json(res.file)
