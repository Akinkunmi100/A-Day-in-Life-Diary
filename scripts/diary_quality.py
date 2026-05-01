import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import pandas as pd

f = r'c:\Users\LUMEN GLOB AL\Documents\TAOF\Thematic Analysis\Merged A day in life diary.xlsx'
df_raw = pd.read_excel(f, engine='openpyxl', header=None)

# Row 1 has the question headers
headers = df_raw.iloc[1].fillna('')

# Data rows start at row 2
df = df_raw.iloc[2:].copy()
df.columns = range(len(df.columns))
df = df.reset_index(drop=True)

print("=== DATA COMPLETENESS PER COLUMN ===")
print(f"{'Col':>4} | {'Non-Null':>8} | {'%Fill':>5} | Question")
print("-" * 100)
for c in df.columns:
    nn = df[c].notna().sum()
    pct = round(nn / len(df) * 100)
    q = str(headers.iloc[c])[:75] if c < len(headers) else ''
    marker = " *** SPARSE" if pct < 30 else ""
    print(f"{c:>4} | {nn:>8} | {pct:>4}% | {q}{marker}")

print()
print("=" * 80)
print("=== RESPONDENT DIARY COUNTS ===")
name_col = 1
names = df[name_col].value_counts()
print(names.to_string())
print(f"\nTotal unique respondents: {len(names)}")
print(f"Expected (7 days x 11 respondents): 77 diaries")
print(f"Actual: {len(df)} diaries")
print(f"Missing: {77 - len(df)} entries")

print()
print("=== DATE COVERAGE ===")
dates = df[2].dropna().apply(lambda x: str(x)[:10])
print(dates.value_counts().sort_index().to_string())

print()
print("=== RESPONDENTS WITH MISSING TIME BLOCK ===")
missing_tb = df[df[3].isna()]
for _, row in missing_tb.iterrows():
    print(f"  ID {int(row[0])} | {row[1]} | {str(row[2])[:10]}")

print()
print("=== KEY QUALITATIVE RICHNESS CHECK ===")
# Morning section columns 26-37 (2-hour check-ins)
# Afternoon section columns 39-68
# Evening section columns 70-82
sections = {
    "Morning routine (cols 5-16)": list(range(5, 17)),
    "Travel (cols 17-25)": list(range(17, 26)),
    "Morning 2hr check-ins (cols 26-37)": list(range(26, 38)),
    "Afternoon activities (cols 39-48)": list(range(39, 49)),
    "Social/Comms (cols 49-59)": list(range(49, 60)),
    "Emotions (cols 60-68)": list(range(60, 69)),
    "Evening (cols 70-82)": list(range(70, 83)),
}

for section, cols in sections.items():
    valid_cols = [c for c in cols if c < len(df.columns)]
    filled = df[valid_cols].notna().sum().sum()
    total = len(df) * len(valid_cols)
    pct = round(filled / total * 100) if total > 0 else 0
    print(f"  {section}: {pct}% filled ({filled}/{total} cells)")

print()
print("=== SAMPLE RICH RESPONSES (Col 42 - How phone helped at work) ===")
samples = df[42].dropna()
for i, v in samples.head(10).items():
    name = df.iloc[i][1]
    print(f"  [{name}]: {str(v)[:120]}")

print()
print("=== SAMPLE RICH RESPONSES (Col 20 - Journey description) ===")
journeys = df[20].dropna()
for i, v in journeys.head(10).items():
    name = df.iloc[i][1]
    print(f"  [{name}]: {str(v)[:120]}")
