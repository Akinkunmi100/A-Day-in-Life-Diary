"""
Full Corpus Coder — 11-Family Codebook (A-K)
Applies segment-level thematic codes to all 65 diary entries.
Uses question-context + keyword-pattern matching.
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import pandas as pd
import re
from collections import Counter

# ============ LOAD DATA ============
f = r'c:\Users\LUMEN GLOB AL\Documents\TAOF\Thematic Analysis\Merged A day in life diary.xlsx'
df_raw = pd.read_excel(f, engine='openpyxl', header=None)
headers = df_raw.iloc[1]
df = df_raw.iloc[2:].copy()
df.columns = range(len(df.columns))
df = df.reset_index(drop=True)

skip_cols = {0,1,2,3,38,69}
trivial = {'nothing','no','yes','nan','none','n/a','nope','','no,','yes,'}

def get_section(c):
    if 4 <= c <= 16: return "Morning Routine"
    if 17 <= c <= 25: return "Travel & Movement"
    if 26 <= c <= 37: return "Morning Check-ins"
    if 39 <= c <= 48: return "Afternoon Activities"
    if 49 <= c <= 59: return "Social & Communication"
    if 60 <= c <= 68: return "Emotional Landscape"
    if 70 <= c <= 82: return "Evening & Relaxation"
    return "Other"

# ============ KEYWORD-BASED CODING RULES ============
# Each rule: (code, keywords_list) — if ANY keyword matches, code applies
# Rules are grouped by question context for precision

def code_response(col, response, question):
    """Apply codes based on column context + keyword patterns."""
    r = response.lower().strip()
    q = question.lower().strip() if question else ''
    codes = []
    
    # Skip trivial
    if r in trivial or r.startswith('nothing') or r == 'no, nothing':
        return ['N00']
    
    # ---- FAMILY A: Work & Productivity ----
    if any(w in r for w in ['work', 'business', 'office', 'shop', 'customer', 'client',
                             'task', 'job', 'trading', 'selling', 'marketing']):
        if col in range(39,49) or col in [15,16,18,40,41]:
            codes.append('A05')
        elif col in [10,11,24,42,43]:
            codes.append('A01')
    if any(w in r for w in ['whatsapp', 'advertis', 'contacts as a method']):
        codes.append('A04')
    if any(w in r for w in ['remote', 'work from home', 'worked remotely']):
        codes.append('A03')
    if any(w in r for w in ['handbill', 'billboard', 'another means', 'available options',
                             'someone else', 'someone elses', 'alternative']):
        if col == 14 or 'without' in q:
            codes.append('I02')
    if any(w in r for w in ['teacher', 'teach', 'insurance', 'designer', 'graphic',
                             'lecturer', 'tailor', 'stylist']):
        if col in [39,40,42,43]:
            codes.append('A06')
    if any(w in r for w in ['submit', 'laptop', 'document', 'editing', 'print']):
        codes.append('A05')
    if 'snap picture' in r or 'photograph' in r:
        codes.append('A01')
    
    # ---- FAMILY B: Network Experience ----
    if any(w in r for w in ['mtn', 'glo', 'airtel', '9mobile', 'etisalat']):
        if col in [54,77]:
            codes.append('B01')
        elif col in [55]:
            codes.append('B01')
    if any(w in r for w in ['network', 'bad network', 'poor network', 'network issue',
                             'network delay']):
        if any(w in q for w in ['stress', 'slow', 'pain', 'challenge']):
            codes.append('B02')
        elif 'could not' in r or 'couldn' in r or 'wanted' in r:
            codes.append('B04')
        else:
            codes.append('B02')
    if any(w in r for w in ['app', 'gtb', 'transfer', 'bank app']):
        if any(w in r for w in ['fail', 'issue', 'bad', 'not working', 'couldn', 'not allow',
                                 'not going through', 'could not']):
            codes.append('B03')
            if 'transfer' in r:
                codes.append('B05')
    if any(w in r for w in ['bundle', 'affordable', 'cheap', 'data is']):
        if col in [55]:
            codes.append('B01')
    if 'main line' in r or 'main and alternative' in r:
        pass  # metadata, not coded
    if any(w in r for w in ['fast network', 'good network', 'smooth', 'network was good']):
        codes.append('B08')
    # Multi-SIM detection
    if col == 77:
        # Evening network — if different from calling network, flag B06
        pass  # handled at entry level below

    # ---- FAMILY C: Mobility & Movement ----
    if any(w in r for w in ['transport', 'traffic', 'road', 'commut', 'hold up',
                             'stressful', 'queuing', 'queueing', 'delayed', 'delay']):
        if any(w in q for w in ['journey', 'stress', 'slow', 'delays', 'challenge']):
            codes.append('C01')
        elif col == 20:
            codes.append('C01')
        elif any(w in q for w in ['stress','slow']):
            codes.append('C01')
    if any(w in r for w in ['bus', 'brt', 'tricycle', 'keke', 'okada', 'bike', 'uber',
                             'bolt', 'ride', 'walk', 'public transportation']):
        if col in [19,20,24,29,32,34,40]:
            codes.append('C02')
    if any(w in r for w in ['curfew', 'couldn\'t go', 'no transport', 'couldn\'t travel',
                             'couldn\'t get']):
        codes.append('C03')
    if col in [23,24] and r not in trivial:
        if 'phone' not in r.lower():
            codes.append('C04')
    if any(w in r for w in ['surulere', 'ojuwoye', 'ogba', 'yaba', 'festac', 'okota',
                             'alagomeji', 'ojuelegba', 'ikeja', 'lekki', 'mushin',
                             'oshodi', 'ikorodu', 'agege', 'maryland']):
        codes.append('C05')
    if any(w in r for w in ['rain', 'weather', 'flood', 'erosion', 'bad road']):
        codes.append('C06')
    if any(w in r for w in ['change', 'delayed of change', 'delayed in change',
                             'my change']):
        if any(w in q for w in ['stress', 'slow']):
            codes.append('C01')

    # ---- FAMILY D: Financial Behaviour ----
    if any(w in r for w in ['food', 'transport', 'data', 'airtime', 'recharge',
                             'bill', 'necessary', 'essential', 'must']):
        if col in [45,46]:
            codes.append('D01')
    if any(w in r for w in ['budget', 'no money', 'poor sales', 'low sales',
                             'little money', 'sales was poor', 'save', 'afford']):
        codes.append('D02')
    if any(w in r for w in ['consider', 'option', 'weigh', 'decide', 'before spending']):
        if col == 47:
            codes.append('D03')
    if any(w in r for w in ['bank app', 'ussd', 'transfer', 'mobile money', 'opay',
                             'palmpay', 'kuda', 'pos']):
        if col in [45,46,47,48]:
            codes.append('D04')
    if any(w in r for w in ['cash', 'gave them cash', 'physical cash']):
        codes.append('D05')
    if any(w in r for w in ['impulse', 'unplanned', 'came up', 'wasn\'t budgeted']):
        codes.append('D06')

    # ---- FAMILY E: Social & Relational Life ----
    if any(w in r for w in ['family', 'parent', 'sibling', 'husband', 'wife', 'kids',
                             'children', 'mum', 'mom', 'dad', 'mother', 'father']):
        if col in [49,50,81] or 'buy' in r or 'for my' in r:
            codes.append('E01')
    if any(w in r for w in ['friend', 'colleague', 'neighbour', 'people', 'gist',
                             'church member']):
        if col in [49,50,51,73,81] or 'visit' in r or 'see' in r:
            codes.append('E02')
    if any(w in r for w in ['community', 'youth', 'opportunit', 'development',
                             'creating', 'young people', 'collective']):
        codes.append('E03')
        codes.append('I05')
    if any(w in r for w in ['call', 'whatsapp', 'sms', 'in-person', 'in person',
                             'face to face', 'visit']):
        if col in [51,52,53]:
            codes.append('E04')
    if any(w in r for w in ['planned', 'scheduled']):
        if col == 57:
            codes.append('E05')
    if any(w in r for w in ['came up', 'random', 'spontaneous', 'just called',
                             'not planned']):
        if col == 57:
            codes.append('E05')
    if any(w in r for w in ['not available', 'couldn\'t reach', 'not on seat',
                             'on leave', 'hacked', 'not around', 'didn\'t meet',
                             'phone was not available']):
        codes.append('E06')

    # ---- FAMILY F: Emotional & Physical Experience ----
    if any(w in r for w in ['happy', 'energetic', 'motivated', 'peaceful', 'relaxed',
                             'good', 'blessed', 'wonderful']):
        if col == 7:
            codes.append('F01')
    if any(w in r for w in ['stress', 'pressur', 'uncomfortable', 'tension']):
        if col in [61,62] or any(w in q for w in ['stress','slow']):
            codes.append('F02')
    if any(w in r for w in ['frustrat', 'block', 'prevent', 'slow me', 'couldn\'t',
                             'not working', 'not active', 'not agile', 'arguing']):
        if col in range(26,38) or col in range(60,69):
            codes.append('F03')
    if any(w in r for w in ['productive', 'satisfied', 'accomplish', 'fruitful',
                             'exciting', 'completed', 'credited', 'sales']):
        if col in [60,59]:
            codes.append('F04')
    if any(w in r for w in ['normal', 'that is how', 'usual', 'always like this']):
        if col == 22:
            codes.append('F05')
    if col in [70,73,75] and any(w in r for w in ['relax', 'rest', 'sleep', 'movie',
                                                    'music', 'game']):
        codes.append('F06')
    if any(w in r for w in ['anxiety', 'anxious', 'worried', 'worry', 'fear',
                             'nervous', 'thoughts how']):
        codes.append('F07')
    if any(w in r for w in ['sick', 'tired', 'fatigue', 'hungry', 'stomach',
                             'headache', 'pain', 'ill', 'sleepy']):
        if col == 7 or 'feel' in q:
            codes.append('F08')

    # ---- FAMILY G: Digital Leisure ----
    if any(w in r for w in ['youtube', 'tiktok', 'netflix', 'video', 'movie',
                             'watching movie']):
        codes.append('G01')
    if any(w in r for w in ['spotify', 'music', 'radio', 'listen to music',
                             'listening to music']):
        codes.append('G02')
    if any(w in r for w in ['game', 'gaming', 'played games']):
        codes.append('G03')
    if any(w in r for w in ['instagram', 'facebook', 'social media', 'scrolling',
                             'browsing', 'tik tok']):
        if col in [75,78] or 'leisure' in q or 'entertainment' in q:
            codes.append('G04')
    if any(w in r for w in ['food', 'dance', 'comedy', 'sport', 'news', 'reading',
                             'motivational']):
        if col == 78:
            codes.append('G05')
    if any(w in r for w in ['escape stress', 'relax', 'stress relief', 'helps me relax',
                             'escape', 'feel good']):
        if col in [79]:
            codes.append('G06')

    # ---- FAMILY H: Routine & Spiritual Life ----
    if any(w in r for w in ['pray', 'prayer', 'devotion', 'morning devotion',
                             'quiet time', 'church', 'mosque', 'imam']):
        if col in [5,6,15,16,26,29,32,35,39]:
            codes.append('H01')
    if any(w in r for w in ['chores', 'bath', 'prepare']):
        if col in [5,6] and 'pray' in r:
            codes.append('H02')
    if any(w in r for w in ['chosen my clothes', 'iron', 'pack for tomorrow',
                             'preparing for tomorrow']):
        if col == 71:
            codes.append('H03')
    if any(w in r for w in ['exercise', 'workout', 'gym', 'jogging', 'walk']):
        if col in [5,6,32,73]:
            codes.append('H04')

    # ---- FAMILY I: Aspirations & Unmet Needs ----
    if any(w in r for w in ['wanted to', 'i wanted', 'couldn\'t', 'could not',
                             'was not available', 'but it was not', 'but couldn']):
        if col in [27,30,33,36,64,67]:
            codes.append('I01')
    if any(w in r for w in ['instead', 'another means', 'find another', 'walk instead',
                             'available options', 'someone else']):
        codes.append('I02')
    if any(w in r for w in ['better network', 'if the network', 'would have been easier']):
        codes.append('I03')
    if any(w in r for w in ['not available', 'not working', 'on leave', 'not ready',
                             'closed', 'out of stock', 'printer', 'hacked']):
        if col in [27,30,33,36,48,64,67]:
            codes.append('I04')
    # I05 already handled under E03

    # ---- FAMILY J: Power & Infrastructure ----
    if any(w in r for w in ['nepa', 'power outage', 'no light', 'power supply',
                             'light', 'electricity']):
        if any(w in r for w in ['no light', 'nepa', 'power outage', 'power supply',
                                 'interrupted']):
            codes.append('J01')
    if any(w in r for w in ['charge', 'battery', 'low batter', 'phone was off',
                             'phone off', 'charge my phone']):
        codes.append('J02')
    if any(w in r for w in ['generator', 'gen', 'power bank', 'alternative charger']):
        codes.append('J03')
    if any(w in r for w in ['iron', 'washing machine', 'blender']):
        if any(w in r for w in ['power', 'light', 'couldn']):
            codes.append('J04')

    # ---- FAMILY K: Domestic Labour ----
    if any(w in r for w in ['cook', 'breakfast', 'lunch', 'dinner', 'meal',
                             'food stuff', 'soup', 'egg', 'coffee']):
        if col in [5,6,15,26,29,32,35,39,70,71]:
            codes.append('K01')
    if any(w in r for w in ['clean', 'house chores', 'chores', 'laundry', 'wash',
                             'fold', 'wardrobe', 'sweep']):
        if col in [5,6,15,26,29,32,35,39,66]:
            codes.append('K02')
    if any(w in r for w in ['kids', 'children', 'child', 'taking care']):
        if 'room' in r or 'care' in r or 'teach' in r:
            codes.append('K03')
    if any(w in r for w in ['market', 'grocery', 'groceries', 'buy food',
                             'food stuff', 'shopping', 'provisions']):
        if col in [29,32,35,39,45]:
            codes.append('K04')
    if any(w in r for w in ['bath', 'dress', 'groom', 'preparing for work',
                             'prepared for work', 'iron']):
        if col in [5,6,16,26]:
            codes.append('K05')

    # Deduplicate
    seen = set()
    unique_codes = []
    for c in codes:
        if c not in seen:
            seen.add(c)
            unique_codes.append(c)
    
    return unique_codes if unique_codes else ['N00']


# ============ PROCESS ALL ENTRIES ============
rows = []
for idx in range(len(df)):
    entry_id = int(df.iloc[idx, 0]) if pd.notna(df.iloc[idx, 0]) else 0
    respondent = str(df.iloc[idx, 1]) if pd.notna(df.iloc[idx, 1]) else 'Unknown'
    date = str(df.iloc[idx, 2])[:10] if pd.notna(df.iloc[idx, 2]) else ''
    tb = str(df.iloc[idx, 3]) if pd.notna(df.iloc[idx, 3]) else 'Unspecified'
    
    # Track calling network and evening network for multi-SIM detection
    call_net = ''
    eve_net = ''
    
    for c in range(4, min(83, len(df.columns))):
        if c in skip_cols:
            continue
        val = df.iloc[idx, c]
        if pd.isna(val) or str(val).strip() == '':
            continue
        
        response = str(val).strip()
        q = str(headers.iloc[c]).strip() if c < len(headers) and pd.notna(headers.iloc[c]) else f"Col {c}"
        section = get_section(c)
        
        # Track networks
        if c == 54:
            call_net = response.lower()
        if c == 77:
            eve_net = response.lower()
        
        # Skip pure Yes/No columns (9,12,13,17,21,23,41,44,72,76,80)
        if c in {9,12,13,17,21,23,41,44,72,76,80}:
            r_lower = response.lower().strip()
            if r_lower in trivial or r_lower in {'yes', 'no'}:
                continue
        
        codes = code_response(c, response, q)
        
        if codes == ['N00']:
            continue  # skip trivial
        
        rows.append({
            'Entry_ID': entry_id,
            'Respondent': respondent,
            'Date': date,
            'Time_Block': tb,
            'Section': section,
            'Col': c,
            'Question': q[:100],
            'Response': response[:250],
            'Code_1': codes[0] if len(codes) > 0 else '',
            'Code_2': codes[1] if len(codes) > 1 else '',
            'Code_3': codes[2] if len(codes) > 2 else '',
            'Code_4': codes[3] if len(codes) > 3 else '',
            'All_Codes': '; '.join(codes),
        })
    
    # Multi-SIM check
    if call_net and eve_net and call_net != eve_net:
        rows.append({
            'Entry_ID': entry_id,
            'Respondent': respondent,
            'Date': date,
            'Time_Block': tb,
            'Section': 'Cross-Entry',
            'Col': 0,
            'Question': 'Multi-SIM Detection',
            'Response': f'Calls: {call_net.upper()} / Evening: {eve_net.upper()}',
            'Code_1': 'B06',
            'Code_2': '',
            'Code_3': '',
            'Code_4': '',
            'All_Codes': 'B06',
        })

out_df = pd.DataFrame(rows)
out_path = r'c:\Users\LUMEN GLOB AL\Documents\TAOF\Thematic Analysis\coded_corpus_full.xlsx'
out_df.to_excel(out_path, index=False, engine='openpyxl')
print(f"Written {len(out_df)} coded segments to {out_path}")

# ============ SUMMARY STATS ============
all_codes = []
for _, row in out_df.iterrows():
    for col in ['Code_1','Code_2','Code_3','Code_4']:
        if row[col]:
            all_codes.append(row[col])

print(f"\nTotal codes applied: {len(all_codes)}")
print(f"Unique codes used: {len(set(all_codes))}")
print(f"Respondents: {out_df['Respondent'].nunique()}")
print(f"Entries coded: {out_df['Entry_ID'].nunique()}")

fam_names = {'A':'Work','B':'Network','C':'Mobility','D':'Financial','E':'Social',
             'F':'Emotional','G':'Leisure','H':'Routine','I':'Aspirations',
             'J':'Power','K':'Domestic','N':'Nothing'}
families = Counter([c[0] for c in all_codes])
print(f"\n{'='*50}")
print(f"CODE FAMILY FREQUENCY (ALL 11 RESPONDENTS)")
print(f"{'='*50}")
for fam in sorted(families.keys()):
    print(f"  {fam} ({fam_names.get(fam,'')}): {families[fam]}")

code_freq = Counter(all_codes)
print(f"\n{'='*50}")
print(f"TOP 25 CODES")
print(f"{'='*50}")
for code, count in code_freq.most_common(25):
    print(f"  {code}: {count}")

# Per-respondent summary
print(f"\n{'='*50}")
print(f"CODES PER RESPONDENT")
print(f"{'='*50}")
for resp in sorted(out_df['Respondent'].unique()):
    resp_df = out_df[out_df['Respondent'] == resp]
    resp_codes = []
    for _, row in resp_df.iterrows():
        for col in ['Code_1','Code_2','Code_3','Code_4']:
            if row[col]:
                resp_codes.append(row[col])
    entries = resp_df['Entry_ID'].nunique()
    print(f"  {resp}: {len(resp_codes)} codes across {entries} entries ({len(resp_codes)/max(entries,1):.0f} codes/entry)")
