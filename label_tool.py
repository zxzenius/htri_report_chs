# Merge en/chs labels in one csv file, for convenience.
from pathlib import Path
import csv
from translate import get_labels


def merge():
    root = Path(r'./Languages')
    files = [root/lang/'LabelTable.txt' for lang in ['English', 'Simplified Chinese']]
    en_labels, chs_labels = [get_labels(file) for file in files]
    out_file = 'labels.csv'
    with open(out_file, 'w', newline='', encoding='utf8') as f:
        fieldnames = ['label', 'en', 'chs']
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for label in en_labels:
            row = {'label': label, 'en': en_labels[label], 'chs': chs_labels.get(label, '')}
            writer.writerow(row)
    # print(len(en_labels))
    print('ok')
