import csv
from collections import defaultdict
from pathlib import Path
import re
from comtypes.client import CreateObject


def get_labels(file):
    db = dict()
    file = Path(file)
    pattern = re.compile(r'^(?P<key>\w+|\+)\s.*\"(?P<value>.+)\"')
    print(f"loading '{file.resolve()}'")
    with open(file, 'r', encoding='utf8') as f:
        for line in f.readlines():
            if match := pattern.match(line):
                key, value = match.groups()
                if key != '+':
                    db[key] = value
                    last_key = key
                else:
                    db[last_key] += value
    return db


def load_labels(file):
    file = Path(file)
    print(f"loading '{file.resolve()}'")
    labels = defaultdict(dict)
    with open(file, 'r', encoding='utf8', newline='') as f:
        reader = csv.DictReader(f)
        for row in reader:
            for lang in list(row)[1:]:
                labels[lang][row['label']] = row[lang]
    return labels


def translate(file, labels: dict, outpath: str = None):
    file = Path(file)
    xl = CreateObject("Excel.Application", dynamic=True)
    xl.DisplayAlerts = False
    print(f"Translating '{file.absolute()}'")
    wb = xl.Workbooks.Open(str(file.resolve()))
    for ws in wb.Worksheets:
        for cell in ws.UsedRange:
            if cell.Value() is not None:
                cell.Value[:] = replace(cell.Value(), labels)
    if outpath:
        if outpath.endswith('/') or outpath.endswith('\\'):
            outpath = Path(outpath)
            if not outpath.exists():
                outpath.mkdir(parents=True)
            outfile = Path(outpath) / file.name
        else:
            outfile = Path(outpath)
        if outfile.exists():
            outfile.unlink()
        print(f"Save to '{outfile.resolve()}'")
        wb.SaveAs(str(outfile.resolve()))
    else:
        wb.Save()
    wb.Close()
    print(f"Finished.")
    xl.Quit()
    xl = None


def replace(string: str, replacements: dict):
    fast = 0
    result = ''
    while fast < len(string):
        if string[fast] == '{':
            slow = fast
            key_start = slow + 1
            key_end = string.find('}', key_start)
            fast = key_end + 1
            if key_end > 0:
                key = string[key_start:key_end].upper()
                if key in replacements:
                    result += replacements[key]
                else:
                    result += string[slow:fast]
        else:
            result += string[fast]
            fast += 1

    return result


def process(path, out=None, lang:str="chs"):
    path = Path(path)
    labels = load_labels('./labels.csv')[lang]
    if path.is_file():
        translate(path, labels, out)
    else:
        files = list(path.glob('*.xls*'))
        for counter, file in enumerate(files, 1):
            print(f'---[{counter}/{len(files)}]---')
            translate(file, labels, out)
