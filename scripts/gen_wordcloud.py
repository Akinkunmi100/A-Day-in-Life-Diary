import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from wordcloud import WordCloud
import pandas as pd

outdir = r'c:\Users\LUMEN GLOB AL\Documents\TAOF\Thematic Analysis'

f = r'c:\Users\LUMEN GLOB AL\Documents\TAOF\Thematic Analysis\Merged A day in life diary.xlsx'
df_raw = pd.read_excel(f, engine='openpyxl', header=None)
df = df_raw.iloc[2:].copy()
df.columns = range(len(df.columns))
df = df.reset_index(drop=True)

pilot = df[df[1] == 'Adekoya Adesola'].copy()

# Collect all text responses
skip_cols = {0,1,2,3,38,69,82}
stop_words = {'nothing','no','yes','nan','none','n/a','nope','the','a','an','to','is','it',
              'i','my','was','and','at','in','for','of','that','but','we','do','did','had',
              'have','been','not','with','this','from','on','about','so','just','are','be',
              'me','would','could','can','will','go','got','get','some','very','also','because',
              'due','there','what','how','when','like','or','if','all','which','by','up','out',
              'they','them','their','its','she','her','he','him','his','we','our','been','being',
              'has','am','were','than','more','most','other','into','over','after','before','then',
              'between','each','few','both','some','such','here','where','why','own','same'}

all_text = []
for _, row in pilot.iterrows():
    for c in range(4, len(row)):
        if c in skip_cols:
            continue
        val = row[c]
        if pd.notna(val):
            text = str(val).strip().lower()
            if text not in {'nothing','no','yes','nan'}:
                all_text.append(text)

combined = ' '.join(all_text)

# Generate word cloud
wc = WordCloud(
    width=1600, height=900,
    background_color='#0D1B2A',
    colormap='cool',
    max_words=120,
    stopwords=stop_words,
    min_font_size=8,
    max_font_size=80,
    prefer_horizontal=0.7,
    relative_scaling=0.5,
    collocations=True,
    margin=15,
).generate(combined)

fig, ax = plt.subplots(1, 1, figsize=(16, 9))
fig.patch.set_facecolor('#0D1B2A')
ax.imshow(wc, interpolation='bilinear')
ax.axis('off')
ax.set_title('Word Cloud: Adekoya Adesola \u2014 7-Day Diary Profile',
             fontsize=18, fontweight='bold', color='white', pad=20)

wc_path = outdir + r'\WordCloud_Adekoya_Adesola.png'
plt.savefig(wc_path, dpi=200, bbox_inches='tight', facecolor='#0D1B2A')
plt.close()
print(f"Word cloud saved to: {wc_path}")
