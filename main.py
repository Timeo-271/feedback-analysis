import os
import pandas as pd
import numpy as np
from collections import Counter
import json
import jieba
import openpyxl.cell._writer

# Import configuration
print("Loading configurations...")
with open(r"_config/setup.json", 'r', encoding='utf-8') as file_config:
    config = json.load(file_config)
years = config['years']
print(f"> Years: {years}")
related_top = config['related_top']
print(f"> Search for related words of the top #{related_top} most frequent words.")
related_words = config['related_words']
print(f"> Search for related words of: {related_words}.")

# Load user-defined dict and filter
print(f"The following words will be added to the dictionary:")
with open(r"_config/dict.txt", 'r', encoding='utf-8') as file_dict:
    for line in file_dict:
        print(f"> {line.strip()}")
jieba.load_userdict(r"_config/dict.txt")
print(f"The following words will be excluded:")
filter = set()
with open(r"_config/filt.txt", 'r', encoding='utf-8') as file_filter:
    for line in file_filter:
        filter.add(line.strip())
        print(f"> {line.strip()}")

# Writer
print("Creating Excel file...")
try:
    writer = pd.ExcelWriter(r"result.xlsx", engine='openpyxl')
except PermissionError:
    print("Close the result.xlsx file first.")
    exit(1)

# Read feedback data
print("Reading data...")
seg_list = []
for year in years:
    print(f"> Year: {year}")
    data_file_name = fr"data/{year}年值班日志记录表.xlsx"
    xls = pd.ExcelFile(data_file_name)
    sheet_count = len(xls.sheet_names)
    for month in range(0, sheet_count):
        data = pd.read_excel(data_file_name, sheet_name=month, usecols=[2], keep_default_na=False)
        data = np.array(data)
        # Cut sentence
        for seq in data:
            seg_list += list(jieba.cut(seq[0], cut_all=False))

# Count and sort
print("Counting word frequency...")
counted = Counter(seg_list)
sorted_counts = sorted(counted.items(), key=lambda x: x[1], reverse=True)

# Output
items = []
counts = []
for item, count in sorted_counts:
    if item in filter:
        continue
    items.append(item)
    counts.append(count)
pd.DataFrame({'词语': items, '计数': counts}).to_excel(writer, sheet_name='ALL', index=False)

# Related word for top words
for word in items[:related_top]:
    print(f"Counting related word frequency for: {word}")
    seg_list = []
    for year in years:
        data_file_name = fr"data/{year}年值班日志记录表.xlsx"
        xls = pd.ExcelFile(data_file_name)
        sheet_count = len(xls.sheet_names)
        for month in range(0, sheet_count):
            data = pd.read_excel(data_file_name, sheet_name=month, usecols=[2], keep_default_na=False)
            data = np.array(data)
            # Cut sentence
            for seq in data:
                seq_list = list(jieba.cut(seq[0], cut_all=False))
                if word in seq_list:
                    seg_list += seq_list

    # Count and sort
    counted = Counter(seg_list)
    sorted_counts = sorted(counted.items(), key=lambda x: x[1], reverse=True)

    # Output
    items = []
    counts = []
    for item, count in sorted_counts:
        if item in filter or item == word:
            continue
        items.append(item)
        counts.append(count)
    pd.DataFrame({'词语': items, '计数': counts}).to_excel(writer, sheet_name=f"top_{word}", index=False)

# Related word for select words
for word in related_words:
    print(f"Counting related word frequency for: {word}")
    seg_list = []
    for year in years:
        data_file_name = fr"data/{year}年值班日志记录表.xlsx"
        xls = pd.ExcelFile(data_file_name)
        sheet_count = len(xls.sheet_names)
        for month in range(0, sheet_count):
            data = pd.read_excel(data_file_name, sheet_name=month, usecols=[2], keep_default_na=False)
            data = np.array(data)
            # Cut sentence
            for seq in data:
                seq_list = list(jieba.cut(seq[0], cut_all=False))
                if word in seq_list:
                    seg_list += seq_list

    # Count and sort
    counted = Counter(seg_list)
    sorted_counts = sorted(counted.items(), key=lambda x: x[1], reverse=True)

    # Output
    items = []
    counts = []
    for item, count in sorted_counts:
        if item in filter or item == word:
            continue
        items.append(item)
        counts.append(count)
    pd.DataFrame({'词语': items, '计数': counts}).to_excel(writer, sheet_name=f"select_{word}", index=False)

writer.close()
print(f"All done! The results are available in: {os.path.abspath("result.xlsx")}")
print("Press any key to exit...")
input()
