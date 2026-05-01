import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import pandas as pd

f = r'c:\Users\LUMEN GLOB AL\Documents\TAOF\Thematic Analysis\Merged A day in life diary.xlsx'
df_raw = pd.read_excel(f, engine='openpyxl', header=None)

# Data starts at row 2 (0-indexed), row 0 is section header, row 1 is question header
df = df_raw.iloc[2:].copy()
df.columns = range(len(df.columns))
df = df.reset_index(drop=True)

id_col = 0
name_col = 1
date_col = 2
timeblock_col = 3
wakeup_col = 4
type_col = 82

print("=== DATASET OVERVIEW ===")
print(f"Total diary entries: {len(df)}")
unique_names = df[name_col].dropna().unique()
print(f"Unique respondents: {len(unique_names)}")
print()

print("=== ALL DIARY ENTRIES (ID | Name | Date | Time Block) ===")
for _, row in df.iterrows():
    rid = str(int(row[id_col])) if pd.notna(row[id_col]) else "?"
    name = str(row[name_col]) if pd.notna(row[name_col]) else "?"
    date = str(row[date_col])[:10] if pd.notna(row[date_col]) else "?"
    tb = str(row[timeblock_col]) if pd.notna(row[timeblock_col]) else "?"
    print(f"  {rid} | {name} | {date} | {tb}")

print()
print("=== WAKE-UP TIMES (Col 4) ===")
wakeups = df[wakeup_col].dropna()
print(wakeups.value_counts().head(15).to_string())

print()
print("=== TIME BLOCK DISTRIBUTION ===")
tb_dist = df[timeblock_col].value_counts()
print(tb_dist.to_string())

print()
print("=== NETWORKS USED IN CALLS (Col 54) ===")
nets = df[54].dropna()
print(nets.value_counts().to_string())

print()
print("=== PHONE USED IMMEDIATELY AFTER WAKING (Col 9) ===")
phone_first = df[9].dropna()
print(phone_first.value_counts().to_string())

print()
print("=== DID YOU LEAVE HOME (Col 17) ===")
left_home = df[17].dropna()
print(left_home.value_counts().to_string())

print()
print("=== TRANSPORT MODE (Col 19) ===")
transport = df[19].dropna()
print(transport.value_counts().to_string())

print()
print("=== ONLINE/OFFLINE TYPE (Col 82) ===")
otype = df[type_col].dropna()
print(otype.value_counts().to_string())

print()
print("=== EVENING RELAXATION CONTENT (Col 78) ===")
econtent = df[78].dropna()
print(econtent.value_counts().head(15).to_string())

print()
print("=== NETWORK FOR EVENING (Col 77) ===")
enet = df[77].dropna()
print(enet.value_counts().to_string())

print()
print("=== STRESSORS / CHALLENGES (Col 61 + 62) ===")
stress = df[61].dropna()
print("Col 61 - What made you stressed:")
print(stress.value_counts().head(20).to_string())
