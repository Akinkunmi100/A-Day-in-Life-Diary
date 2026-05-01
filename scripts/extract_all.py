"""Extract full diary text for each respondent into separate files."""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import pandas as pd

outdir = r'C:\Users\LUMEN GLOB AL\.gemini\antigravity\scratch'
f = r'c:\Users\LUMEN GLOB AL\Documents\TAOF\Thematic Analysis\Merged A day in life diary.xlsx'
df_raw = pd.read_excel(f, engine='openpyxl', header=None)
headers = df_raw.iloc[1]
df = df_raw.iloc[2:].copy()
df.columns = range(len(df.columns))
df = df.reset_index(drop=True)

coded = pd.read_excel(r'c:\Users\LUMEN GLOB AL\Documents\TAOF\Thematic Analysis\coded_corpus_full.xlsx', engine='openpyxl')

skip_cols = {0,1,2,3,38,69}
trivial = {'nothing','no','yes','nan','none','n/a','nope',''}

for resp_name in df[1].unique():
    if resp_name == 'Adekoya Adesola':
        continue  # already done
    
    resp_df = df[df[1] == resp_name].reset_index(drop=True)
    safe = resp_name.replace(' ', '_')
    
    with open(f'{outdir}\\data_{safe}.txt', 'w', encoding='utf-8') as fout:
        fout.write(f"RESPONDENT: {resp_name}\n")
        fout.write(f"ENTRIES: {len(resp_df)}\n\n")
        
        # Code profile summary
        rc = coded[coded['Respondent'] == resp_name]
        from collections import Counter
        all_codes = []
        for _, row in rc.iterrows():
            for c in ['Code_1','Code_2','Code_3','Code_4']:
                if pd.notna(row[c]) and row[c]:
                    all_codes.append(row[c])
        fam_freq = Counter([c[0] for c in all_codes])
        code_freq = Counter(all_codes)
        fout.write("CODE PROFILE:\n")
        fn = {'A':'Work','B':'Network','C':'Mobility','D':'Financial','E':'Social',
              'F':'Emotional','G':'Leisure','H':'Routine','I':'Aspirations','J':'Power','K':'Domestic'}
        for f_letter, cnt in fam_freq.most_common():
            fout.write(f"  {f_letter} ({fn.get(f_letter,'')}): {cnt}\n")
        fout.write(f"\nTOP 10 CODES: {code_freq.most_common(10)}\n\n")
        
        for entry_idx, (_, entry_row) in enumerate(resp_df.iterrows()):
            eid = int(entry_row[0]) if pd.notna(entry_row[0]) else 0
            date = str(entry_row[2])[:10] if pd.notna(entry_row[2]) else '?'
            tb = str(entry_row[3]) if pd.notna(entry_row[3]) else '?'
            fout.write(f"{'='*80}\n")
            fout.write(f"ENTRY {entry_idx+1}/{len(resp_df)} | ID: {eid} | Date: {date} | Time Block: {tb}\n")
            fout.write(f"{'='*80}\n")
            
            for c in range(4, min(83, len(entry_row))):
                if c in skip_cols: continue
                val = entry_row[c]
                if pd.isna(val): continue
                v = str(val).strip()
                if not v: continue
                
                q = str(headers.iloc[c]).strip() if c < len(headers) and pd.notna(headers.iloc[c]) else f"Col {c}"
                q = q.replace('\u25cf','*').replace('\u25cb','o').replace('\t',' ')[:80]
                fout.write(f"  Q[{c:2d}] {q}\n")
                fout.write(f"         -> {v}\n\n")
            fout.write("\n")
    
    print(f"Extracted: {resp_name} ({len(resp_df)} entries)")

print("\nAll respondent data extracted!")
