"""
Generate Word Clouds + Journey Maps for all 11 respondents.
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
from wordcloud import WordCloud
import pandas as pd
from collections import Counter

outdir = r'c:\Users\LUMEN GLOB AL\Documents\TAOF\Thematic Analysis'

# Load raw data
f = outdir + r'\Merged A day in life diary.xlsx'
df_raw = pd.read_excel(f, engine='openpyxl', header=None)
headers = df_raw.iloc[1]
df = df_raw.iloc[2:].copy()
df.columns = range(len(df.columns))
df = df.reset_index(drop=True)

# Load coded corpus
coded = pd.read_excel(outdir + r'\coded_corpus_full.xlsx', engine='openpyxl')

skip_cols = {0,1,2,3,38,69,82}
trivial_words = {'nothing','no','yes','nan','none','n/a','nope',''}
stop_words = {'nothing','no','yes','nan','none','n/a','nope','','the','a','an','to','is','it',
              'i','my','was','and','at','in','for','of','that','but','we','do','did','had',
              'have','been','not','with','this','from','on','about','so','just','are','be',
              'me','would','could','can','will','go','got','get','some','very','also','because',
              'due','there','what','how','when','like','or','if','all','which','by','up','out',
              'they','them','their','its','she','her','he','him','his','we','our','been','being',
              'has','am','were','than','more','most','other','into','over','after','before','then'}

respondents = df[1].unique()
print(f"Processing {len(respondents)} respondents...\n")

for resp_name in respondents:
    print(f"--- {resp_name} ---")
    resp_df = df[df[1] == resp_name]
    resp_coded = coded[coded['Respondent'] == resp_name]
    safe_name = resp_name.replace(' ', '_')
    
    # ===== WORD CLOUD =====
    all_text = []
    for _, row in resp_df.iterrows():
        for c in range(4, min(83, len(row))):
            if c in skip_cols: continue
            val = row[c]
            if pd.notna(val):
                text = str(val).strip().lower()
                if text not in trivial_words and len(text) > 2:
                    all_text.append(text)
    
    combined = ' '.join(all_text)
    if len(combined) < 20:
        print(f"  Skipping word cloud - insufficient text")
        continue
    
    wc = WordCloud(width=1600, height=900, background_color='#0D1B2A', colormap='cool',
                   max_words=100, stopwords=stop_words, min_font_size=8, max_font_size=70,
                   prefer_horizontal=0.7, relative_scaling=0.5, collocations=True, margin=12
                  ).generate(combined)
    
    fig, ax = plt.subplots(1, 1, figsize=(16, 9))
    fig.patch.set_facecolor('#0D1B2A')
    ax.imshow(wc, interpolation='bilinear')
    ax.axis('off')
    ax.set_title(f'Word Cloud: {resp_name}', fontsize=18, fontweight='bold', color='white', pad=20)
    plt.savefig(f'{outdir}\\WordCloud_{safe_name}.png', dpi=150, bbox_inches='tight', facecolor='#0D1B2A')
    plt.close()
    print(f"  Word cloud saved")
    
    # ===== JOURNEY MAP =====
    # Extract data per time block
    def get_responses(cols):
        texts = []
        for _, row in resp_df.iterrows():
            for c in cols:
                if c < len(row) and pd.notna(row[c]):
                    t = str(row[c]).strip()
                    if t.lower() not in trivial_words and len(t) > 2:
                        texts.append(t)
        return texts
    
    # Get morning feeling for emotion score
    feelings = []
    for _, row in resp_df.iterrows():
        if 7 < len(row) and pd.notna(row[7]):
            feelings.append(str(row[7]).lower())
    
    pos_words = ['happy','energetic','relaxed','peaceful','good','motivated','blessed','healthy']
    neg_words = ['sick','tired','sleepy','cold','stressed','anxious']
    
    morning_score = 0
    for f_text in feelings:
        if any(w in f_text for w in pos_words): morning_score += 1
        if any(w in f_text for w in neg_words): morning_score -= 1
    morning_emo = max(-3, min(3, morning_score))
    
    # Get stress mentions per section
    stress_cols_morning = [28,31,34,37]
    stress_cols_afternoon = [61,62,65,68]
    
    def count_stress(cols):
        count = 0
        total = 0
        for _, row in resp_df.iterrows():
            for c in cols:
                if c < len(row) and pd.notna(row[c]):
                    t = str(row[c]).strip().lower()
                    total += 1
                    if t not in trivial_words and t != 'no, nothing':
                        count += 1
        return count, total
    
    morn_stress, morn_total = count_stress(stress_cols_morning)
    aftn_stress, aftn_total = count_stress(stress_cols_afternoon)
    
    # Get top activities
    def summarize(texts, n=4):
        if not texts: return ['—']
        words = Counter()
        for t in texts:
            for w in t.lower().split():
                if w not in stop_words and len(w) > 2:
                    words[w] += 1
        return [w for w, _ in words.most_common(n)]
    
    morning_acts = summarize(get_responses([5,6,15,16,26]), 4)
    commute_acts = summarize(get_responses([19,20,24]), 3)
    work_acts = summarize(get_responses([39,40,42]), 4)
    social_acts = summarize(get_responses([49,50,51]), 3)
    evening_acts = summarize(get_responses([70,73,75,78]), 4)
    
    # Get networks
    call_nets = Counter()
    eve_nets = Counter()
    for _, row in resp_df.iterrows():
        if 54 < len(row) and pd.notna(row[54]):
            call_nets[str(row[54]).upper()] += 1
        if 77 < len(row) and pd.notna(row[77]):
            eve_nets[str(row[77]).upper()] += 1
    
    top_call = call_nets.most_common(1)[0][0] if call_nets else '?'
    top_eve = eve_nets.most_common(1)[0][0] if eve_nets else '?'
    
    # Build emotion curve
    commute_emo = -2.0 if morn_stress > 0 else 0.0
    late_morn_emo = -1.0 if morn_stress > 1 else 0.5
    afternoon_emo = -1.5 if aftn_stress > 0 else 0.5
    late_aftn_emo = 0.5
    evening_emo = 2.0  # universally positive
    
    emotions = [morning_emo, commute_emo, late_morn_emo, afternoon_emo, late_aftn_emo, evening_emo]
    
    stages = ['Pre-Dawn\n05:00-07:00', 'Morning\nCommute\n07:00-09:00', 'Late Morning\n09:00-12:00',
              'Afternoon\n12:00-16:00', 'Late\nAfternoon\n16:00-18:00', 'Evening\n18:00-22:00']
    x_pos = np.arange(len(stages))
    
    actions = [
        '\n'.join(morning_acts[:4]),
        '\n'.join(commute_acts[:3]),
        '\n'.join(work_acts[:4]),
        '\n'.join(social_acts[:3]),
        'Planned calls\nErrands',
        '\n'.join(evening_acts[:4]),
    ]
    
    touchpoints = [
        f'{top_call}: messages',
        f'{top_call}: calls',
        f'{top_call}: work',
        f'{top_call}: voice',
        f'{top_call}: calls',
        f'{top_eve}: data',
    ]
    
    # Get top stressors
    stress_texts = get_responses(stress_cols_morning + stress_cols_afternoon)
    stress_words = summarize(stress_texts, 3) if stress_texts else ['—']
    
    pains = [
        '\n'.join([f.capitalize() for f in feelings[:2]]) if feelings else '',
        '\n'.join(stress_words[:2]),
        '\n'.join(stress_words[1:3]) if len(stress_words) > 1 else '',
        '\n'.join(stress_words[:2]),
        '',
        '',
    ]
    
    # Draw journey map
    fig, axes = plt.subplots(2, 1, figsize=(18, 11), gridspec_kw={'height_ratios': [3, 1]})
    fig.patch.set_facecolor('#0D1B2A')
    
    ax = axes[0]
    ax.set_facecolor('#0D1B2A')
    
    for i, stage in enumerate(stages):
        color = '#1B2838' if i % 2 == 0 else '#1E3A5F'
        rect = mpatches.FancyBboxPatch((i-0.4, 2.8), 0.8, 0.8, boxstyle="round,pad=0.05",
                                         facecolor=color, edgecolor='#2E86AB', linewidth=1.5)
        ax.add_patch(rect)
        ax.text(i, 3.2, stage, ha='center', va='center', fontsize=8, fontweight='bold', color='white')
        
        rect2 = mpatches.FancyBboxPatch((i-0.45, 1.6), 0.9, 1.0, boxstyle="round,pad=0.03",
                                          facecolor='#162B44', edgecolor='#3A7CA5', linewidth=0.8)
        ax.add_patch(rect2)
        ax.text(i, 2.1, actions[i], ha='center', va='center', fontsize=7, color='#B8D4E3')
        
        rect3 = mpatches.FancyBboxPatch((i-0.45, 0.4), 0.9, 1.0, boxstyle="round,pad=0.03",
                                          facecolor='#0A2239', edgecolor='#2E86AB', linewidth=0.8)
        ax.add_patch(rect3)
        ax.text(i, 0.9, touchpoints[i], ha='center', va='center', fontsize=7, color='#7EC8E3')
        
        if pains[i]:
            rect4 = mpatches.FancyBboxPatch((i-0.45, -0.8), 0.9, 1.0, boxstyle="round,pad=0.03",
                                              facecolor='#3D0C11', edgecolor='#CC3300', linewidth=0.8)
            ax.add_patch(rect4)
            ax.text(i, -0.3, pains[i], ha='center', va='center', fontsize=7, color='#FF6B6B')
    
    ax.text(-0.75, 3.2, 'TIME\nBLOCK', ha='center', va='center', fontsize=8, fontweight='bold', color='#2E86AB')
    ax.text(-0.75, 2.1, 'ACTIONS', ha='center', va='center', fontsize=8, fontweight='bold', color='#2E86AB')
    ax.text(-0.75, 0.9, 'PHONE &\nNETWORK', ha='center', va='center', fontsize=8, fontweight='bold', color='#2E86AB')
    ax.text(-0.75, -0.3, 'PAIN\nPOINTS', ha='center', va='center', fontsize=8, fontweight='bold', color='#CC3300')
    ax.set_xlim(-1, 6)
    ax.set_ylim(-1.2, 4)
    ax.axis('off')
    
    # Get top code family for subtitle
    resp_codes = []
    for _, row in resp_coded.iterrows():
        for c in ['Code_1','Code_2','Code_3','Code_4']:
            if pd.notna(row[c]) and row[c]:
                resp_codes.append(row[c])
    fam_freq = Counter([c[0] for c in resp_codes])
    top_fam = fam_freq.most_common(1)[0][0] if fam_freq else '?'
    fam_names = {'A':'Work-Driven','B':'Network-Focused','C':'Mobility-Stressed',
                 'D':'Finance-Anxious','E':'Socially-Connected','F':'Emotionally-Complex',
                 'G':'Digitally-Engaged','H':'Ritual-Anchored','I':'Aspiration-Driven',
                 'J':'Power-Constrained','K':'Domestically-Centred'}
    archetype = fam_names.get(top_fam, '')
    
    ax.set_title(f'Customer Journey Map: {resp_name}',
                 fontsize=16, fontweight='bold', color='white', pad=20)
    
    # Emotion curve
    ax2 = axes[1]
    ax2.set_facecolor('#0D1B2A')
    ax2.fill_between(x_pos, emotions, 0, where=[e >= 0 for e in emotions],
                      alpha=0.3, color='#4ECDC4', interpolate=True)
    ax2.fill_between(x_pos, emotions, 0, where=[e < 0 for e in emotions],
                      alpha=0.3, color='#FF6B6B', interpolate=True)
    ax2.plot(x_pos, emotions, '-o', color='white', linewidth=2.5, markersize=10, zorder=5)
    ax2.axhline(y=0, color='#334455', linestyle='--', linewidth=0.8)
    ax2.set_xticks(x_pos)
    ax2.set_xticklabels([s.replace('\n', ' ') for s in stages], fontsize=8, color='#AAAAAA')
    ax2.set_ylabel('Emotional State', fontsize=10, color='#2E86AB')
    ax2.set_ylim(-3, 3)
    ax2.tick_params(colors='#666666')
    ax2.spines['bottom'].set_color('#334455')
    ax2.spines['left'].set_color('#334455')
    ax2.spines['top'].set_visible(False)
    ax2.spines['right'].set_visible(False)
    ax2.set_title('Emotion Curve', fontsize=11, color='#AAAAAA', pad=10)
    
    plt.tight_layout()
    plt.savefig(f'{outdir}\\Journey_Map_{safe_name}.png', dpi=150, bbox_inches='tight', facecolor='#0D1B2A')
    plt.close()
    print(f"  Journey map saved")

print(f"\nAll visualisations complete!")
