import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import pandas as pd

f = r'c:\Users\LUMEN GLOB AL\Documents\TAOF\Thematic Analysis\Merged A day in life diary.xlsx'
df_raw = pd.read_excel(f, engine='openpyxl', header=None)
headers = df_raw.iloc[1]
df = df_raw.iloc[2:].copy()
df.columns = range(len(df.columns))
df = df.reset_index(drop=True)

pilot = df[df[1] == 'Adekoya Adesola'].copy().reset_index(drop=True)

# Manual coding for Adekoya Adesola - all 7 entries
# Each entry: list of (col, response_snippet, codes, notes)

coding_data = []
skip_cols = {0,1,2,3,38,69}

# Section mapping
def get_section(c):
    if 4 <= c <= 16: return "Morning Routine"
    if 17 <= c <= 25: return "Travel & Movement"
    if 26 <= c <= 37: return "Morning Check-ins"
    if 39 <= c <= 48: return "Afternoon Activities"
    if 49 <= c <= 59: return "Social & Communication"
    if 60 <= c <= 68: return "Emotional Landscape"
    if 70 <= c <= 82: return "Evening & Relaxation"
    return "Other"

# Coding rules applied manually per response
# Format: {entry_idx: {col: (codes_list, note)}}
codings = {
    # ENTRY 1 (60003, Apr 22)
    0: {
        4: ([], "Metadata: 07:00"),
        5: (["H01","H02","K05"], "Prayer + structured routine + bathing"),
        7: (["F08"], "Sick - physical health barrier"),
        8: (["A06"], "Work schedule determines morning - professional identity"),
        10: (["A01"], "Phone for messages/info - business tool"),
        11: (["A01","E04"], "Customer communication - phone as business tool"),
        14: (["I02"], "Workaround: use available options"),
        15: (["K02","A05"], "Chores + work as priorities"),
        16: (["K05","A05"], "Preparing for work"),
        18: (["A05"], "Destination: work"),
        19: (["C02"], "Public transportation"),
        20: (["C01"], "Stressful journey - noise and bad road"),
        24: (["C04","A01"], "Phone while mobile for information"),
        26: (["K05"], "Prepared for work"),
        27: (["I01","E06"], "Wanted to see friend but couldn't"),
        28: (["C01","C06"], "Transport + bad weather as stressors"),
        29: (["E01","K04"], "Picking up package for friend"),
        30: (["I01","I04"], "Store didn't have item - service gap"),
        31: (["C01","B02"], "Transport and network slowed down"),
        32: (["E01","K04","C05"], "Shopping for parent at Ojuwoye"),
        33: (["I01","C03"], "Curfew prevented going out - mobility constraint"),
        34: (["B03","B05"], "GTB transfer failed - bad network"),
        35: (["A05","C05"], "Conference at Surulere"),
        36: (["E06","I01"], "Friend's phone not available"),
        37: (["C06","C01"], "Bad weather + transportation"),
        39: (["H01","K02"], "Morning devotion + house chores"),
        40: (["A05"], "Work required most effort"),
        42: (["A04"], "WhatsApp for sending information"),
        43: (["A01"], "Interviewing - phone essential for work"),
        45: (["E01","D01","K04"], "Buying food for parent - necessity + family obligation"),
        46: (["E02","A05"], "Talking about business with friend"),
        47: (["D02","D03"], "Budget consideration before spending"),
        48: (["B03","B05","D04"], "Bank app failed due to network"),
        49: (["E01"], "Family interaction"),
        50: (["E03","I05"], "Discussed making money + opportunities for youth"),
        51: (["E04"], "Calls"),
        52: (["E04"], "Convenient method"),
        54: (["B01"], "GLO - network choice"),
        55: (["B01"], "Bundle is affordable - loyalty rationale"),
        57: (["E05"], "Planned conversation - exciting"),
        60: (["F04","D02"], "Getting credited = satisfaction"),
        61: (["F02","C01"], "Transportation stress"),
        62: (["F07"], "Anxiety"),
        63: (["C03"], "Going out"),
        64: (["I01"], "Couldn't update diary"),
        65: (["C03"], "No transportation"),
        66: (["G02"], "Listening to music"),
        67: (["B04","I04"], "Ride app not working - network blocked goal"),
        68: (["D02","F02"], "Slow sales - financial stress"),
        70: (["A05"], "Closed from work"),
        71: (["H03"], "Chose clothes for tomorrow"),
        73: (["G03"], "Played games"),
        75: (["G01"], "Movies"),
        77: (["B06"], "MTN for evening (GLO for calls) - multi-SIM"),
        78: (["G05"], "Food content"),
        79: (["G06"], "Escape stress"),
        81: (["E02","G03"], "Played games with others"),
    },
    # ENTRY 2 (60009, Apr 21)
    1: {
        5: (["H01","H02","K02"], "Prayer + chores"),
        7: (["F01"], "Energetic - positive morning"),
        8: (["A06","F05"], "Previous day + work schedule"),
        10: (["A01"], "Check messages"),
        11: (["E02","E04"], "Chat with friends"),
        14: (["I02"], "Use available options / someone else's phone"),
        15: (["A05"], "Work/business priority"),
        16: (["K05"], "Preparing for work"),
        18: (["A05"], "Work"),
        19: (["C02"], "Public transportation"),
        20: (["C01"], "Bad road"),
        24: (["G03","A01","C04"], "Games + info while mobile"),
        26: (["H02"], "Just waking up - routine"),
        27: (["I01","D02"], "Wanted to buy something"),
        28: (["C01","C06","B02"], "Transport + weather + network"),
        29: (["A05","E02"], "Meeting someone in business"),
        30: (["I01","C01","E06"], "Driver delayed, person not around"),
        31: (["C01","F02"], "Queueing for bus, change issues"),
        32: (["H04"], "Workout around street"),
        33: (["B04","I04"], "Network blocked sneaker purchase"),
        34: (["F03","I04"], "Workers not active - service gap"),
        35: (["A05"], "Documenting events"),
        36: (["E06","I01"], "Conversation failed - lack of understanding"),
        37: (["C01","B02"], "Transportation and network"),
        70: (["K01"], "Made dinner"),
        71: (["H03"], "Chose clothes"),
        73: (["G03"], "Played games"),
        75: (["G01"], "Movies"),
        77: (["B01"], "GLO for evening"),
        78: (["G05"], "Food content"),
        79: (["G06"], "Escape stress"),
        81: (["E02"], "Discussed with family/friends"),
    },
    # ENTRY 3 (60019, Apr 20)
    2: {
        5: (["H01","K05"], "Prayer + bath"),
        7: (["F08"], "Sick"),
        8: (["A06"], "Work schedule"),
        10: (["A01"], "Check messages"),
        11: (["E02","E04"], "Chat with friends"),
        14: (["I02"], "Use available options"),
        15: (["K02","A05"], "Chores + work"),
        16: (["K05"], "Preparing for work"),
        18: (["A05"], "Work"),
        19: (["C02"], "Public transportation"),
        20: (["C01","C06"], "Bad road - drainage system"),
        24: (["A01","C04"], "Phone while mobile to call someone"),
        26: (["K05"], "Prepared for work"),
        27: (["I01"], "Went to fetch water"),
        28: (["C01","C06"], "Transportation + bad weather"),
        29: (["E02"], "Meeting someone for information"),
        30: (["I01","I04"], "Pharmacy didn't have drugs"),
        31: (["C01","B02"], "Transportation + bad network"),
        32: (["K04","C05"], "Getting items at market"),
        33: (["I01","I04","E06"], "Airtel customer care on leave"),
        34: (["F03","I04"], "Mall workers not agile"),
        35: (["H01","A05"], "Documenting church events"),
        36: (["E06","I01"], "Imam not available"),
        37: (["F03"], "False information about project"),
        39: (["K04","C05","E02"], "Shopping at Ojuwoye + met classmate"),
        40: (["A05"], "Work + customer schedule"),
        42: (["A04","A01"], "WhatsApp advertising new product"),
        43: (["A05"], "Offloading goods"),
        45: (["D01","K04"], "Food stuff + recharge"),
        46: (["A04","A05"], "Advertising new product"),
        47: (["D02","D03"], "Budget consideration"),
        48: (["B03","D05"], "App failed, paid cash instead"),
        49: (["E01"], "Family"),
        50: (["E03","I05"], "Career opportunities discussion"),
        51: (["E04"], "Calls"),
        54: (["B01"], "MTN"),
        55: (["B01"], "Bundle affordable"),
        57: (["E05","A05"], "Planned conversation about fabric production"),
        60: (["F04","D02"], "Sales = satisfaction"),
        61: (["D02","F02"], "No money = stress"),
        62: (["F07"], "Anxiety"),
        66: (["G02"], "Listening to music"),
        67: (["I01","I04"], "Printer not working at cyber cafe"),
        68: (["B02"], "Poor network"),
        70: (["K01"], "Made dinner"),
        71: (["H03"], "Chose clothes"),
        73: (["E02","E04"], "Communicated with family/friends"),
        75: (["G01"], "Movies"),
        77: (["B06"], "MTN for evening"),
        78: (["G05"], "Food content"),
        79: (["G06"], "Escape stress"),
        81: (["E02"], "Discussed"),
    },
    # ENTRY 4 (60029, Apr 19)
    3: {
        5: (["H01","K02"], "Prayer + chores"),
        7: (["F01"], "Energetic"),
        8: (["A06"], "Work schedule"),
        10: (["A01"], "Check messages"),
        11: (["E02"], "Chat with friends"),
        14: (["I02"], "Use available options"),
        15: (["A05","K02"], "Work + chores"),
        16: (["K05"], "Preparing for work"),
        18: (["A05"], "Work"),
        19: (["C02"], "Public transportation"),
        20: (["C01"], "Bad road"),
        24: (["A01","C04"], "Phone to call someone while mobile"),
        26: (["K04"], "Went out to get something"),
        27: (["I01","H04"], "Wanted gym"),
        28: (["C01"], "Transportation"),
        29: (["A05"], "Meeting customer"),
        30: (["B04","I01"], "Couldn't book ride due to network"),
        31: (["C06","B02"], "Weather + network"),
        32: (["A05"], "Documenting life"),
        33: (["I01","I04"], "Trader not available"),
        34: (["C06","F03"], "Weather + workers"),
        35: (["E02"], "Composing song for friend"),
        36: (["I01","I04"], "Tailor clothes not ready"),
        37: (["C01","C06"], "Transportation + weather + erosion"),
        39: (["H01"], "Preparing for church"),
        40: (["C01","C02"], "Ordering transportation"),
        42: (["A01","E02"], "Info + communicate with friend"),
        43: (["I02","A01"], "Would use friend's phone"),
        45: (["D01","K04"], "Food stuff"),
        46: (["E01"], "Pass info to siblings"),
        47: (["D02","D03"], "Budget first"),
        48: (["B03","D05"], "App failed, went home"),
        49: (["E01"], "Family"),
        50: (["E03","I05"], "Money + opportunities"),
        51: (["E04"], "Calls"),
        54: (["B01"], "GLO"),
        55: (["B01"], "Bundle affordable"),
        57: (["E05","I05"], "Planned - future prospects"),
        60: (["F04"], "Getting credited"),
        61: (["F02","A05"], "Work stress"),
        62: (["F07"], "Anxiety"),
        63: (["G02"], "Listening to music"),
        64: (["I01"], "Couldn't do more work"),
        65: (["C01"], "Traffic"),
        66: (["K02"], "Chores"),
        67: (["I01","I04"], "Cafeteria food not available"),
        68: (["E02","F02"], "Family/friends as stressor"),
        70: (["G01"], "Watching movies"),
        71: (["H03"], "Chose clothes"),
        73: (["G03"], "Played games"),
        75: (["G02"], "Listening to music"),
        77: (["B01"], "GLO"),
        78: (["G05"], "Dance content"),
        79: (["G06"], "Escape stress"),
        81: (["E02"], "Discussed"),
    },
    # ENTRY 5 (60038, Apr 18)
    4: {
        5: (["H01","K02"], "Prayer + chores"),
        7: (["F01"], "Energetic"),
        8: (["A06"], "Work schedule"),
        10: (["A01"], "Check messages"),
        11: (["E02"], "Chat with friends"),
        14: (["I02"], "Use someone else's phone"),
        15: (["A05","K02"], "Work + chores"),
        16: (["K05"], "Preparing for work"),
        18: (["A05"], "Work"),
        19: (["C02"], "Public transportation"),
        20: (["C01"], "Stressful"),
        24: (["A01","C04"], "Call someone while mobile"),
        26: (["K05"], "Prepared for work"),
        27: (["I01","D02"], "Wanted to buy something"),
        28: (["C01","C06"], "Transport + bad weather"),
        29: (["A05"], "Making a delivery"),
        30: (["E06","I01"], "Friend not available"),
        31: (["C06","B02"], "Weather + network"),
        32: (["E02","C05"], "Visit friend at Festac"),
        33: (["B04","I04"], "Uber app not working"),
        34: (["C01","D02"], "Change from driver + queueing"),
        35: (["G01"], "Preparing for movie show"),
        36: (["I01","E01"], "Couldn't buy drug for friend"),
        37: (["F03","C01"], "Workers + change issues"),
        70: (["E02"], "Went out"),
        71: (["H03"], "Chose clothes"),
        73: (["G03"], "Played games"),
        75: (["G02"], "Listening to music"),
        77: (["B01"], "GLO"),
        78: (["G05"], "Sports content"),
        79: (["G06"], "Escape stress"),
        81: (["E02","G03"], "Played games together"),
    },
    # ENTRY 6 (60047, Apr 17)
    5: {
        5: (["H01","K02"], "Prayer + chores"),
        7: (["F01"], "Happy + energetic"),
        8: (["A06"], "Work schedule"),
        10: (["A01"], "Check messages"),
        11: (["E02"], "Chat with friends"),
        14: (["I02"], "Use someone else's phone"),
        15: (["F07"], "Thoughts about how day will go"),
        16: (["K05"], "Preparing for work"),
        18: (["A05"], "Work"),
        19: (["C02"], "Public transportation"),
        20: (["C01"], "Stressful"),
        24: (["A01","C04"], "Check information while mobile"),
        26: (["K05"], "Prepared for work"),
        27: (["I01","F08"], "Wanted to eat but couldn't"),
        28: (["B02"], "Network"),
        29: (["K04"], "Getting custard"),
        30: (["E06","B04"], "Couldn't reach loved one - phone + network"),
        31: (["B02"], "Network"),
        32: (["H04"], "Gym + bath + nap"),
        33: (["E06","B04"], "Couldn't chat on TikTok"),
        34: (["C06","C01"], "Weather + transportation"),
        35: (["E02","F04"], "People came over - exciting"),
        36: (["E06","I04"], "Facebook hacked by scammers"),
        37: (["F03","C01"], "Arguing + transportation"),
        39: (["A05","K02"], "Work + chores"),
        40: (["A05"], "Work"),
        42: (["A01"], "Snap pictures + get info"),
        43: (["E02","I02"], "Would visit friend instead"),
        45: (["D01","J01"], "Food + NEPA bill"),
        46: (["E02"], "Listening to people"),
        47: (["D02","D03"], "Budget"),
        48: (["B03","B05"], "GTB transfer failed - bad network"),
        49: (["E01"], "Family"),
        50: (["E03","I05"], "Productivity + opportunities"),
        51: (["E04"], "Calls"),
        54: (["B01"], "GLO"),
        55: (["B01"], "Fast network"),
        57: (["E05","E03"], "Planned - social issues discussion"),
        60: (["F04"], "Completing tasks"),
        61: (["F02","C01"], "Transportation"),
        62: (["F07"], "Anxiety"),
        63: (["A05"], "Work"),
        64: (["I01","E01"], "Wanted to visit sick friend"),
        65: (["C03"], "No transportation"),
        66: (["C01"], "Travelling"),
        67: (["B03","B05","I04"], "Transfer failed at pharmacy"),
        68: (["C01"], "Traffic"),
        70: (["G01"], "Watching movies"),
        71: (["H03"], "Chose clothes"),
        73: (["G02"], "Listen to music"),
        75: (["E02","E04"], "Chat with friends"),
        77: (["B01"], "GLO"),
        78: (["G05"], "Food content"),
        79: (["G06"], "Relax"),
        81: (["E02","G03"], "Played games"),
    },
    # ENTRY 7 (60057, Apr 16)
    6: {
        5: (["H01","K02"], "Prayer + chores"),
        7: (["F08"], "Sick"),
        8: (["A06"], "Work schedule"),
        10: (["A01"], "Check messages"),
        11: (["A01","E02"], "Customer + friends"),
        14: (["I02"], "Use available options"),
        15: (["F07"], "Thoughts about the day"),
        16: (["K05"], "Preparing for work"),
        18: (["A05"], "Work"),
        19: (["C02"], "Public transportation"),
        20: (["C01"], "Stressful"),
        24: (["A01","C04"], "Check information"),
        26: (["A01"], "Checked phone for messages"),
        27: (["I01","E06"], "Wanted to see friend"),
        28: (["C06"], "Bad weather"),
        29: (["K04"], "Getting food"),
        30: (["I01","C06","C03"], "Hospital blocked by rain"),
        31: (["C01","D02"], "Delayed change by driver"),
        32: (["K04"], "Getting groceries"),
        33: (["B04","I04"], "Couldn't watch movie - network"),
        34: (["C01","D02"], "Transportation + delayed change"),
        35: (["K04"], "Getting groceries"),
        36: (["I01","I04"], "No water in environment"),
        37: (["C01","D02"], "Transportation + delayed change"),
        70: (["E02"], "Went out"),
        71: (["H03"], "Chose clothes"),
        73: (["E02"], "Went out with friends"),
        75: (["G01"], "Movies"),
        77: (["B01"], "GLO"),
        78: (["G05"], "Food content"),
        79: (["G06","F06"], "Relax"),
        81: (["E02"], "Discussed"),
    },
}

# Build output rows
rows = []
for entry_idx, (_, entry_row) in enumerate(pilot.iterrows()):
    entry_id = int(entry_row[0])
    date = str(entry_row[2])[:10]
    tb = str(entry_row[3]) if pd.notna(entry_row[3]) else 'Unspecified'
    
    entry_codes = codings.get(entry_idx, {})
    
    for c in sorted(entry_codes.keys()):
        codes, note = entry_codes[c]
        if not codes:
            continue
        
        q = str(headers.iloc[c]).strip() if c < len(headers) and pd.notna(headers.iloc[c]) else f"Col {c}"
        q = q.replace('\u25cf','*').replace('\u25cb','o').replace('\t',' ')
        
        response = str(entry_row[c]).strip() if pd.notna(entry_row[c]) else ''
        section = get_section(c)
        
        rows.append({
            'Entry_ID': entry_id,
            'Respondent': 'Adekoya Adesola',
            'Date': date,
            'Time_Block': tb,
            'Section': section,
            'Col': c,
            'Question': q[:100],
            'Response': response[:200],
            'Code_1': codes[0] if len(codes) > 0 else '',
            'Code_2': codes[1] if len(codes) > 1 else '',
            'Code_3': codes[2] if len(codes) > 2 else '',
            'Coding_Notes': note,
        })

out_df = pd.DataFrame(rows)
out_path = r'c:\Users\LUMEN GLOB AL\Documents\TAOF\Thematic Analysis\pilot_coded_corpus.xlsx'
out_df.to_excel(out_path, index=False, engine='openpyxl')
print(f"Written {len(out_df)} coded segments to {out_path}")

# Summary stats
all_codes = []
for _, row in out_df.iterrows():
    for col in ['Code_1','Code_2','Code_3']:
        if row[col]:
            all_codes.append(row[col])

print(f"\nTotal codes applied: {len(all_codes)}")
print(f"Unique codes used: {len(set(all_codes))}")

# Family frequency
from collections import Counter
families = Counter([c[0] for c in all_codes])
print(f"\n=== CODE FAMILY FREQUENCY ===")
for fam in sorted(families.keys()):
    fam_names = {'A':'Work','B':'Network','C':'Mobility','D':'Financial','E':'Social',
                 'F':'Emotional','G':'Leisure','H':'Routine','I':'Aspirations',
                 'J':'Power','K':'Domestic','N':'Nothing'}
    print(f"  {fam} ({fam_names.get(fam,'')}): {families[fam]}")

# Code-level frequency
code_freq = Counter(all_codes)
print(f"\n=== TOP 20 CODES ===")
for code, count in code_freq.most_common(20):
    print(f"  {code}: {count}")
