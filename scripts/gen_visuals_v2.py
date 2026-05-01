import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
from wordcloud import WordCloud
import pandas as pd

outdir = r'c:\Users\LUMEN GLOB AL\Documents\TAOF\Thematic Analysis'

# ============================================================
# HELPER FUNCTIONS
# ============================================================
def make_journey_map(name, archetype, stages, actions, touchpoints, pains, emotions, emo_labels, days_label):
    fig, axes = plt.subplots(2, 1, figsize=(18, 12), gridspec_kw={'height_ratios': [3, 1]})
    fig.patch.set_facecolor('#0D1B2A')

    ax = axes[0]
    ax.set_facecolor('#0D1B2A')
    x_pos = np.arange(len(stages))

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

    ax.set_xlim(-1, len(stages))
    ax.set_ylim(-1.2, 4)
    ax.axis('off')
    ax.set_title(f'Customer Journey Map: {name} \u2014 "{archetype}"',
                 fontsize=16, fontweight='bold', color='white', pad=20)

    # Emotion curve
    ax2 = axes[1]
    ax2.set_facecolor('#0D1B2A')
    ax2.fill_between(x_pos, emotions, 0, where=[e >= 0 for e in emotions],
                      alpha=0.3, color='#4ECDC4', interpolate=True)
    ax2.fill_between(x_pos, emotions, 0, where=[e < 0 for e in emotions],
                      alpha=0.3, color='#FF6B6B', interpolate=True)
    ax2.plot(x_pos, emotions, '-o', color='white', linewidth=2.5, markersize=10, zorder=5)

    for i, (x, y, label) in enumerate(zip(x_pos, emotions, emo_labels)):
        color = '#4ECDC4' if y >= 0 else '#FF6B6B'
        ax2.annotate(label, (x, y), textcoords="offset points",
                      xytext=(0, 18 if y >= 0 else -28), ha='center',
                      fontsize=7, fontweight='bold', color=color)

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
    ax2.set_title(f'Emotion Curve (averaged across {days_label})', fontsize=11, color='#AAAAAA', pad=10)

    plt.tight_layout()
    safe_name = name.replace(' ', '_')
    path = f'{outdir}\\Journey_Map_{safe_name}.png'
    plt.savefig(path, dpi=200, bbox_inches='tight', facecolor='#0D1B2A')
    plt.close()
    print(f"Journey map saved: {path}")


def make_wordcloud(name, respondent_filter, days_label):
    f = r'c:\Users\LUMEN GLOB AL\Documents\TAOF\Thematic Analysis\Merged A day in life diary.xlsx'
    df_raw = pd.read_excel(f, engine='openpyxl', header=None)
    df = df_raw.iloc[2:].copy()
    df.columns = range(len(df.columns))
    df = df.reset_index(drop=True)

    pilot = df[df[1] == respondent_filter].copy()

    skip_cols = {0,1,2,3,38,69,82}
    stop_words = {'nothing','no','yes','nan','none','n/a','nope','the','a','an','to','is','it',
                  'i','my','was','and','at','in','for','of','that','but','we','do','did','had',
                  'have','been','not','with','this','from','on','about','so','just','are','be',
                  'me','would','could','can','will','go','got','get','some','very','also','because',
                  'due','there','what','how','when','like','or','if','all','which','by','up','out',
                  'they','them','their','its','she','her','he','him','his','we','our','been','being',
                  'has','am','were','than','more','most','other','into','over','after','before','then',
                  'between','each','few','both','some','such','here','where','why','own','same',
                  'tho','didn','doesn','don','didn\'t','doesn\'t','don\'t','today'}

    all_text = []
    for _, row in pilot.iterrows():
        for c in range(4, len(row)):
            if c in skip_cols:
                continue
            val = row[c]
            if pd.notna(val):
                text = str(val).strip().lower()
                if text not in {'nothing','no','yes','nan','nope','nothing tho','nothing for now'}:
                    all_text.append(text)

    combined = ' '.join(all_text)
    if not combined.strip():
        print(f"WARNING: No text data for {name}, skipping word cloud")
        return

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
    ax.set_title(f'Word Cloud: {name} \u2014 {days_label} Diary Profile',
                 fontsize=18, fontweight='bold', color='white', pad=20)

    safe_name = name.replace(' ', '_')
    path = f'{outdir}\\WordCloud_{safe_name}.png'
    plt.savefig(path, dpi=200, bbox_inches='tight', facecolor='#0D1B2A')
    plt.close()
    print(f"Word cloud saved: {path}")


# ============================================================
# 1. ADEOLA OLOWOLAGBA
# ============================================================
make_journey_map(
    name='Adeola Olowolagba',
    archetype='The Besieged Trader',
    stages=['Pre-Dawn\n05:30-07:00', 'Morning\n07:00-10:00', 'Late Morning\n10:00-12:00',
            'Afternoon\n12:00-16:00', 'Late\nAfternoon\n16:00-18:00', 'Evening\n18:00-22:00'],
    actions=[
        'Prayer\nHouse chores\nCooking\nChildren to school',
        'Walk/bus\nto shop\nTraffic delays\nArrive late',
        'Open shop\nWait for\ncustomers\nPost business',
        'Attend customers\nOnline orders\nConfirm payments\nWhatsApp ads',
        'Close shop\nVisit sister\nChurch visit\nBuy food',
        'Cooking\nWashing\nComedy/movies\nSleep'
    ],
    touchpoints=[
        'Airtel: check\ntime & messages',
        'Airtel: calls\nMTN: browse',
        'WhatsApp:\nonline customers\nSocial media posts',
        'Airtel: customer\ncalls\nBank transfer\nalerts',
        'MTN: internet\nbrowsing',
        'MTN: comedy\nmovies\nWhatsApp chat'
    ],
    pains=[
        'No light\nCan\'t charge\nphone',
        'Traffic\ntied me down',
        'No sales\nWaiting for\ncustomers',
        'Phone hanging\nNo money for\nairtime',
        'No network\nto browse',
        'Can\'t eat well\nNeed to\nspend little'
    ],
    emotions=[0.5, -1.5, -2.0, -1.0, 0.5, 1.0],
    emo_labels=['Happy:\nPrayer grounds', 'STRESS:\nTraffic & late', 'ANXIETY:\nNo sales',
                'Frustrated:\nPhone issues', 'Relief:\nSocial visits', 'ESCAPE:\nComedy & sleep'],
    days_label='6 days'
)

make_wordcloud('Adeola Olowolagba', 'Adeola Olowolagba', '6-Day')


# ============================================================
# 2. ATOLAGBE FLORA YEMI
# ============================================================
make_journey_map(
    name='Atolagbe Flora Yemi',
    archetype='The Weather-Dependent Matriarch',
    stages=['Pre-Dawn\n06:30-08:00', 'Morning\n08:00-10:00', 'Late Morning\n10:00-12:00',
            'Afternoon\n12:00-16:00', 'Late\nAfternoon\n16:00-18:00', 'Evening\n18:00-22:00'],
    actions=[
        'Prayer (6/6)\nHouse chores\nCook breakfast\nCall husband',
        'Walk to shop\nwith kids\nMuddy roads\nafter rain',
        'Open shop\nAttend customers\nVerify payments\nPatiently wait',
        'Confirm\ntransactions\nMake sales\nTravel (Shagamu)',
        'Close shop\nWalk home\nNo transport\navailable',
        'Make dinner\nFamily movies\nChat friends\nCharge phone'
    ],
    touchpoints=[
        'Glo: call\nhusband\nCheck up daily',
        'Phone off\nduring walk\nwith kids',
        'Glo: verify\ntransaction\nalerts',
        'MTN: browse\nfor business\nWhatsApp chat',
        'Glo: calls\nto friends',
        'MTN: TikTok\ncomedy\nbrowsing'
    ],
    pains=[
        'No light\nCan\'t charge\nCan\'t cook fast',
        'Muddy roads\nRain reduces\npatronage',
        'Poor network\nSlow sales\nFew customers',
        'Low battery\nCan\'t record\nPhone hanging',
        'No transport\noptions',
        'Charging\nanxiety for\ntomorrow'
    ],
    emotions=[1.0, -1.5, -1.0, -1.5, -0.5, 1.5],
    emo_labels=['Peaceful:\nPrayer & family', 'STRESS:\nMuddy roads', 'Patient:\nWaiting',
                'ANXIETY:\nLow battery', 'Tired:\nNo transport', 'WARM:\nFamily time'],
    days_label='6 days'
)

make_wordcloud('Atolagbe Flora Yemi', 'Atolagbe Flora Yemi', '6-Day')


# ============================================================
# 3. ADEKUNLE ADEPETUN
# ============================================================
make_journey_map(
    name='Adekunle Adepetun',
    archetype='The Power-Obsessed Shopkeeper',
    stages=['Pre-Dawn\n04:30-06:00', 'Morning\n06:00-09:00', 'Late Morning\n09:00-12:00',
            'Afternoon\n12:00-16:00', 'Late\nAfternoon\n16:00-18:00', 'Evening\n18:00-22:00'],
    actions=[
        'Turn off alarm\nPrayer (6/6)\nHouse chores\nBath',
        'Walk to shop\n(at residence)\nOpen shop\nAwait customers',
        'Sell market\nCalculate\nprices\nAttend customers',
        'Council dispute\nSettle permit\nCall chairman\nBike for charge',
        'Close shop\nWork continues\nMissed meals',
        'Sleep (4/6)\nTikTok (3/6)\nBrowse internet\nCharge phone'
    ],
    touchpoints=[
        'MTN: alarm\nTorch light\nCheck time',
        'MTN: calls\nto customers\nCalculator app',
        'MTN: WhatsApp\ncalls (cheap)\nBusiness chat',
        'MTN: call\nCDA chairman\nCrisis mgmt',
        'Phone dead\nIn-person only\nNeighbour chat',
        'MTN: TikTok\nbrowsing\nStress escape'
    ],
    pains=[
        'Phone not\ncharged\nNo power',
        'Low sales\nRain affects\nshop',
        'NEPA: no light\nFlat battery\nAll day',
        'Govt council\ndispute\n3 time blocks',
        'Hungry\nCan\'t eat well\nSpend little',
        'Faulty\npower bank\nNo light'
    ],
    emotions=[1.0, -0.5, -2.0, -2.5, -1.5, 1.0],
    emo_labels=['Happy:\nPrayer starts', 'Hopeful:\nMake sales', 'CRISIS:\nFlat battery',
                'PANIC:\nCouncil + no power', 'Depleted:\nHungry', 'ESCAPE:\nTikTok relief'],
    days_label='6 days'
)

make_wordcloud('Adekunle Adepetun', 'Adekunle Adepetun', '6-Day')


# ============================================================
# 4. DARAMOLA SOLOMON
# ============================================================
make_journey_map(
    name='Daramola Solomon',
    archetype='The Mobile Professional',
    stages=['Morning\n07:00-08:00', 'Commute\n08:00-09:00', 'Late Morning\n09:00-12:00',
            'Afternoon\n12:00-16:00', 'Late\nAfternoon\n16:00-18:00', 'Evening\n18:00-22:00'],
    actions=[
        'Bath\nCup of coffee\nNo phone yet\nPrepare for work',
        'Public transport\n10 min drive\nCheck phone\nfor info',
        'Arrive office\nClient requests\nPolicy renewals\nMarketing runs',
        'Visit brokers\nFollow up tasks\nNew tasks\ninterrupt',
        'Complete\noutstanding\ntasks\nWrap up day',
        'Make dinner\nLaundry/chores\nWatch videos\nMovies online'
    ],
    touchpoints=[
        'Phone unused\nCoffee first\nNo scrolling',
        'MTN: check\ncontacts\nCall planning',
        'MTN: calls\nWhatsApp\nBusiness chat',
        'MTN: customer\ncalls (3 min)\nData for work',
        'MTN: follow-up\nmessages\nSchedule mgmt',
        'MTN: streaming\nvideos online\nRelax + escape'
    ],
    pains=[
        'Woke up\ntired',
        'Rain delay\nTraffic',
        'New tasks\ninterrupt\nplanned work',
        'Cash payment\nfriction\nNo digital option',
        'Transport\nstress',
        'Charging\nanxiety'
    ],
    emotions=[-0.5, -1.0, 0.5, -1.5, -0.5, 1.5],
    emo_labels=['Tired:\nFatigued start', 'STRESS:\nRain + traffic', 'Productive:\nClient work',
                'FRUSTRATED:\nCash + interrupts', 'Pressured:\nTransport', 'REGULATED:\nVideos = calm'],
    days_label='1 day (snapshot)'
)

make_wordcloud('Daramola Solomon', 'Daramola Solomon', '1-Day')


# ============================================================
# 5. MARY AJIFOWOWE OLUWASEUN
# ============================================================
make_journey_map(
    name='Mary Ajifowowe Oluwaseun',
    archetype='The Stoic Commuter',
    stages=['Pre-Dawn\n05:00-06:00', 'Morning\nCommute\n06:00-09:00', 'Late Morning\n09:00-12:00',
            'Afternoon\n12:00-16:00', 'Late\nAfternoon\n16:00-18:00', 'Evening\n18:00-22:00'],
    actions=[
        'Prayer (2/5)\nCheck phone\nPlay music\nMood sets day',
        'Keke to Ogba\nBus to Ojota\nBike when late\nBuy colleague food',
        'Arrive work\nArrange table\nEat breakfast\nStart duties',
        'Work & work\nNothing exciting\nConfirm info\nOnline orders',
        'Still working\nWant to sleep\nWant market\nCan\'t leave',
        'Went out (rare)\nBrowse internet\nNothing (2/5)\nNo relaxation'
    ],
    touchpoints=[
        'MTN: torch\nMusic player\nPhone check',
        'MTN: bank\ntransfer\nin transit',
        'MTN: online\norders &\nenquiries',
        'MTN: calls\nto customers\nBusiness talk',
        'MTN: calls\nto friends\nJust gisting',
        'MTN: browsing\n(1/5 days)\nRarely online'
    ],
    pains=[
        'Woke up sick\n(1/5 days)\nPhone uncharged',
        'Rain started\nas leaving\nLate = scolding',
        '',
        'Work monotony\nBoredom\nNo excitement',
        'Can\'t sleep\nCan\'t shop\nCan\'t do laundry',
        'No relaxation\ntime (2/5)\nRarely sees friends'
    ],
    emotions=[1.0, -1.0, 0.0, -0.5, -1.5, 0.5],
    emo_labels=['Happy:\n3/5 mornings', 'STRESS:\nRain + lateness', 'Neutral:\nRoutine work',
                'MONOTONY:\nWork & work', 'COMPRESSED:\nBlocked desires', 'Rare joy:\nFriends'],
    days_label='5 days'
)

make_wordcloud('Mary Ajifowowe Oluwaseun', 'Mary Ajifowowe Oluwaseun', '5-Day')

print("\n=== ALL 5 JOURNEY MAPS + WORD CLOUDS GENERATED ===")
