import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import pandas as pd
from collections import Counter

df = pd.read_excel(r'c:\Users\LUMEN GLOB AL\Documents\TAOF\Thematic Analysis\coded_corpus_full.xlsx', engine='openpyxl')

print(f"Total rows: {len(df)}")
print(f"Columns: {list(df.columns)}")

# 1. DEAD CODE ANALYSIS
all_possible = ['A01','A02','A03','A04','A05','A06',
                'B01','B02','B03','B04','B05','B06','B07','B08',
                'C01','C02','C03','C04','C05','C06',
                'D01','D02','D03','D04','D05','D06',
                'E01','E02','E03','E04','E05','E06',
                'F01','F02','F03','F04','F05','F06','F07','F08',
                'G01','G02','G03','G04','G05','G06',
                'H01','H02','H03','H04',
                'I01','I02','I03','I04','I05',
                'J01','J02','J03','J04',
                'K01','K02','K03','K04','K05']

used_codes = set()
for col in ['Code_1','Code_2','Code_3','Code_4']:
    used_codes.update(df[col].dropna().unique())
used_codes.discard('')

dead = set(all_possible) - used_codes
print(f"\n{'='*60}")
print(f"1. DEAD CODES (never applied): {len(dead)}")
print(f"{'='*60}")
for c in sorted(dead):
    print(f"   {c}")

# 2. SAMPLE CODED SEGMENTS - spot check 5 random per family
print(f"\n{'='*60}")
print(f"2. SPOT CHECK - Sample coded segments")
print(f"{'='*60}")

for fam_letter in 'ABCDEFGHIJK':
    fam_rows = df[df['Code_1'].str.startswith(fam_letter, na=False)]
    if len(fam_rows) == 0:
        continue
    sample = fam_rows.sample(min(3, len(fam_rows)), random_state=42)
    print(f"\n  --- Family {fam_letter} (n={len(fam_rows)} primary) ---")
    for _, row in sample.iterrows():
        resp = str(row['Response'])[:80]
        codes = str(row['All_Codes'])
        print(f"    [{row['Respondent'][:15]}] Q{row['Col']}: \"{resp}\"")
        print(f"       -> Codes: {codes}")

# 3. CONSISTENCY CHECK - same question, different respondents
print(f"\n{'='*60}")
print(f"3. CROSS-RESPONDENT CONSISTENCY (same Q, different coding?)")
print(f"{'='*60}")

# Check Q5/6 (first things) - should mostly get H01/H02/K02
q5 = df[df['Col'] == 5]
print(f"\n  Q5 (First things you did) - {len(q5)} responses:")
q5_codes = Counter()
for _, row in q5.iterrows():
    for c in ['Code_1','Code_2','Code_3','Code_4']:
        if pd.notna(row[c]) and row[c]:
            q5_codes[row[c]] += 1
for code, cnt in q5_codes.most_common(10):
    print(f"    {code}: {cnt}")

# Check Q20 (journey description) - should mostly get C01
q20 = df[df['Col'] == 20]
print(f"\n  Q20 (Describe journey) - {len(q20)} responses:")
q20_codes = Counter()
for _, row in q20.iterrows():
    for c in ['Code_1','Code_2','Code_3','Code_4']:
        if pd.notna(row[c]) and row[c]:
            q20_codes[row[c]] += 1
for code, cnt in q20_codes.most_common(10):
    print(f"    {code}: {cnt}")

# Check Q54 (network for calling)
q54 = df[df['Col'] == 54]
print(f"\n  Q54 (Network for calling) - {len(q54)} responses:")
q54_vals = Counter(q54['Response'].str.upper())
for val, cnt in q54_vals.most_common():
    print(f"    {val}: {cnt}")

# Check Q77 (evening network)
q77 = df[df['Col'] == 77]
print(f"\n  Q77 (Evening network) - {len(q77)} responses:")
q77_vals = Counter(q77['Response'].str.upper())
for val, cnt in q77_vals.most_common():
    print(f"    {val}: {cnt}")

# 4. MULTI-SIM DETECTION
print(f"\n{'='*60}")
print(f"4. MULTI-SIM BEHAVIOUR (B06)")
print(f"{'='*60}")
b06 = df[df['Code_1'] == 'B06']
for _, row in b06.iterrows():
    print(f"  {row['Respondent']}: {row['Response']}")

# 5. LOW-DENSITY RESPONDENTS
print(f"\n{'='*60}")
print(f"5. LOW-DENSITY RESPONDENT CHECK")
print(f"{'='*60}")
for resp in df['Respondent'].unique():
    resp_df = df[df['Respondent'] == resp]
    entries = resp_df['Entry_ID'].nunique()
    codes_total = 0
    for _, row in resp_df.iterrows():
        for c in ['Code_1','Code_2','Code_3','Code_4']:
            if pd.notna(row[c]) and row[c]:
                codes_total += 1
    density = codes_total / max(entries, 1)
    flag = " ⚠️ LOW" if density < 20 else ""
    print(f"  {resp}: {density:.0f} codes/entry ({entries} entries){flag}")

# 6. RESPONSES WITH NO CODES (should be empty after filtering)
no_code = df[df['Code_1'].isna() | (df['Code_1'] == '')]
print(f"\n{'='*60}")
print(f"6. UNTAGGED SEGMENTS: {len(no_code)}")
print(f"{'='*60}")
if len(no_code) > 0:
    for _, row in no_code.head(5).iterrows():
        print(f"  [{row['Respondent'][:15]}] Q{row['Col']}: \"{str(row['Response'])[:80]}\"")

# 7. ADEKOYA COMPARISON: pilot manual vs automated
print(f"\n{'='*60}")
print(f"7. PILOT COMPARISON: Adekoya manual vs automated")
print(f"{'='*60}")
adekoya = df[df['Respondent'] == 'Adekoya Adesola']
auto_codes = Counter()
for _, row in adekoya.iterrows():
    for c in ['Code_1','Code_2','Code_3','Code_4']:
        if pd.notna(row[c]) and row[c]:
            auto_codes[row[c]] += 1
print(f"  Automated coding: {sum(auto_codes.values())} total codes")
print(f"  Manual pilot: 457 total codes")
print(f"  Difference: {sum(auto_codes.values()) - 457} ({(sum(auto_codes.values())/457 - 1)*100:+.0f}%)")
print(f"\n  Top 10 codes (automated):")
for code, cnt in auto_codes.most_common(10):
    print(f"    {code}: {cnt}")
