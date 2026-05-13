import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
from wordcloud import WordCloud
import pandas as pd
from collections import Counter
import re
import random

outdir = r'c:\Users\LUMEN GLOB AL\Documents\TAOF\Thematic Analysis'

# ============================================================
# 1. LOAD DATA — ALL RESPONDENTS
# ============================================================
f = r'c:\Users\LUMEN GLOB AL\Documents\TAOF\Thematic Analysis\Merged A day in life diary.xlsx'
df_raw = pd.read_excel(f, engine='openpyxl', header=None)
df = df_raw.iloc[2:].copy()
df.columns = range(len(df.columns))
df = df.reset_index(drop=True)

respondents = df[1].dropna().unique().tolist()
print(f"Found {len(respondents)} respondents: {respondents}")

# ============================================================
# 2. COLUMN FILTERING — only open-ended text responses
# ============================================================
# Skip: metadata (0-3), blank separators (38, 69), type label (82),
#        wake-up time col (4), duplicate col (6),
#        yes/no cols (9, 12, 17, 21, 22, 23, 25, 41, 44, 54, 56, 59, 72, 76, 77, 80)
skip_cols = {
    0, 1, 2, 3,         # metadata: serial, name, date, time block
    4,                   # wake-up time (datetime, not text)
    6,                   # exact duplicate of col 5
    9,                   # "Did you use phone?" — yes/no
    12,                  # "Was phone charged?" — yes/no
    17,                  # "Did you leave home?" — yes/no
    21,                  # "Any delays?" — yes/no
    22,                  # "Normal or unusual?" — short
    23,                  # "Use phone while moving?" — yes/no
    25,                  # "Could you do this without phone?" — yes/no
    38,                  # blank separator column
    41,                  # "Did phone help in work?" — yes/no
    44,                  # "Did you spend money?" — yes/no
    54,                  # "What network?" — single word (MTN/Glo/Airtel)
    56,                  # "Is it your main line?" — yes/no
    59,                  # "Was it fruitful?" — yes/no
    69,                  # blank separator column
    72,                  # "Did you have time to relax?" — yes/no
    76,                  # "Did you use internet?" — yes/no
    77,                  # "What network?" — single word
    80,                  # "Spend time with family/friends?" — yes/no
    82,                  # "Type" — Online/Offline metadata
}

# Comprehensive stop words
stop_words = {
    'nothing', 'no', 'yes', 'nan', 'none', 'n/a', 'nope', 'nil',
    'the', 'a', 'an', 'to', 'is', 'it', 'i', 'my', 'was', 'and', 'at',
    'in', 'for', 'of', 'that', 'but', 'we', 'do', 'did', 'had', 'have',
    'been', 'not', 'with', 'this', 'from', 'on', 'about', 'so', 'just',
    'are', 'be', 'me', 'would', 'could', 'can', 'will', 'go', 'got',
    'get', 'some', 'very', 'also', 'because', 'due', 'there', 'what',
    'how', 'when', 'like', 'or', 'if', 'all', 'which', 'by', 'up', 'out',
    'they', 'them', 'their', 'its', 'she', 'her', 'he', 'him', 'his',
    'we', 'our', 'been', 'being', 'has', 'am', 'were', 'than', 'more',
    'most', 'other', 'into', 'over', 'after', 'before', 'then', 'between',
    'each', 'few', 'both', 'some', 'such', 'here', 'where', 'why', 'own',
    'same', 'tho', 'didn', 'doesn', 'don', "didn't", "doesn't", "don't",
    'today', 'went', 'still', 'back', 'really', 'around', 'came', 'come',
    'make', 'made', 'much', 'thing', 'things', 'even', 'though', 'well',
    'one', 'two', 'three', 'first', 'last', 'time', 'day', 'days',
    'morning', 'afternoon', 'evening', 'night', 'nothing tho',
    'nothing for now', 'too', 'way', 'any', 'who', 'whom', 'you', 'your',
}

# Throwaway full-cell responses
throwaway = {
    'nothing', 'no', 'yes', 'nan', 'nope', 'nothing tho',
    'nothing for now', 'nil', 'na', 'normal', 'yes it is',
    'no due to movie', 'some times', 'online', 'offline',
}

# ============================================================
# 3. EXTRACT ALL TEXT — only from rich open-ended columns
# ============================================================
all_words = []

for _, row in df.iterrows():
    for c in range(4, len(row)):
        if c in skip_cols:
            continue
        val = row[c]
        if pd.notna(val):
            text = str(val).strip().lower()
            if text in throwaway or len(text) < 4:
                continue
            # Tokenise and filter
            tokens = re.findall(r"[a-z']+", text)
            for tok in tokens:
                if tok not in stop_words and len(tok) > 2:
                    all_words.append(tok)

word_freq = Counter(all_words)
print(f"\nTotal unique words: {len(word_freq)}")
print(f"Total word occurrences: {sum(word_freq.values())}")
print(f"\nTop 30 words:")
for word, count in word_freq.most_common(30):
    print(f"  {word:20s} -> {count}")

# ============================================================
# 4. COLOUR FUNCTION — VARIETY OF DISTINCT COLOURS
# ============================================================
# A rich palette of visually distinct, vibrant colours.
# Higher frequency words get picked from warm/hot end,
# lower frequency words from cool/muted end,
# but within each band the specific colour is randomised
# for visual variety.

# Define distinct colour bands by frequency tier
hot_colors = [       # dominant words — bold, attention-grabbing
    '#FF2D55',       # hot pink
    '#FF3B30',       # vivid red
    '#FF6B6B',       # coral
    '#E91E63',       # magenta
    '#FF4081',       # pink accent
]

warm_colors = [      # high-frequency — warm, energetic
    '#FF9500',       # orange
    '#FFCC00',       # gold
    '#FFD740',       # amber
    '#F9A825',       # deep yellow
    '#FF8C00',       # dark orange
]

mid_colors = [       # moderate — fresh, vibrant
    '#00D992',       # electric green
    '#4CD964',       # lime green
    '#66BB6A',       # green
    '#26C6DA',       # cyan
    '#00BCD4',       # teal
]

cool_colors = [      # low-frequency — cool, calm
    '#5AC8FA',       # sky blue
    '#007AFF',       # royal blue
    '#42A5F5',       # material blue
    '#7986CB',       # indigo
    '#64B5F6',       # light blue
]

rare_colors = [      # rare words — muted, pastel
    '#AF52DE',       # purple
    '#BA68C8',       # lavender
    '#9575CD',       # deep purple
    '#7E57C2',       # violet
    '#CE93D8',       # light purple
]

max_freq = max(word_freq.values())

def variety_color_func(word, font_size, position, orientation, random_state=None, **kwargs):
    """Assign a colour from a distinct palette band based on word frequency."""
    freq = word_freq.get(word.lower(), 1)
    ratio = freq / max_freq

    if ratio >= 0.6:
        palette = hot_colors
    elif ratio >= 0.35:
        palette = warm_colors
    elif ratio >= 0.18:
        palette = mid_colors
    elif ratio >= 0.08:
        palette = cool_colors
    else:
        palette = rare_colors

    return random.choice(palette)


# ============================================================
# 5. GENERATE THE COMBINED WORD CLOUD
# ============================================================
wc = WordCloud(
    width=2400,
    height=1400,
    background_color='#0D1B2A',
    max_words=200,
    stopwords=stop_words,
    min_font_size=10,
    max_font_size=120,
    prefer_horizontal=0.65,
    relative_scaling=0.55,
    collocations=True,
    margin=12,
    color_func=variety_color_func,
).generate_from_frequencies(word_freq)

# ============================================================
# 6. RENDER
# ============================================================
fig = plt.figure(figsize=(24, 15), facecolor='#0D1B2A')

# Main word cloud
ax_cloud = fig.add_axes([0.03, 0.06, 0.94, 0.84])
ax_cloud.set_facecolor('#0D1B2A')
ax_cloud.imshow(wc, interpolation='bilinear')
ax_cloud.axis('off')

# Title
ax_cloud.set_title(
    'Combined Word Cloud — All Respondents · Lagos Telecom Ethnography',
    fontsize=26, fontweight='bold', color='white', pad=28,
    fontfamily='sans-serif'
)

# Subtitle
fig.text(
    0.5, 0.93,
    f'{len(respondents)} respondents · {sum(word_freq.values()):,} words · {len(word_freq):,} unique terms',
    ha='center', va='center', fontsize=14, color='#7B8D9E',
    fontfamily='sans-serif', fontstyle='italic'
)

# Colour legend — small swatches showing the 5 tiers
legend_y = 0.025
tier_labels = [
    ('Dominant', hot_colors[0]),
    ('High', warm_colors[0]),
    ('Moderate', mid_colors[0]),
    ('Low', cool_colors[0]),
    ('Rare', rare_colors[0]),
]
for i, (label, color) in enumerate(tier_labels):
    x = 0.28 + i * 0.11
    fig.patches.append(plt.Rectangle((x, legend_y), 0.015, 0.015,
                                      facecolor=color, edgecolor='#334455',
                                      linewidth=0.5, transform=fig.transFigure))
    fig.text(x + 0.02, legend_y + 0.007, label, fontsize=10, color='#AABBCC',
             va='center', fontfamily='sans-serif')

fig.text(0.5, 0.005, 'Word Frequency Intensity',
         ha='center', va='center', fontsize=10, color='#556677',
         fontfamily='sans-serif')

# Save
out_path = outdir + r'\WordCloud_Combined_All_Respondents.png'
plt.savefig(out_path, dpi=200, bbox_inches='tight', facecolor='#0D1B2A')
plt.close()
print(f"\n✓ Combined word cloud saved to: {out_path}")
