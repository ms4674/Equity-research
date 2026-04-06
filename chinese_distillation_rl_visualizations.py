"""
Chinese Model Distillation & RL Analysis — Visualization Script

Generates publication-quality charts comparing reinforcement learning approaches
and distillation strategies across Chinese open-source, Western open-source,
and Western proprietary LLMs.

Outputs: PNG charts saved to working directory.
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch
import matplotlib.patheffects as path_effects
import seaborn as sns

plt.rcParams.update({
    'font.size': 11,
    'font.family': 'DejaVu Sans',
    'axes.titlesize': 14,
    'axes.titleweight': 'bold',
    'axes.labelsize': 12,
    'figure.facecolor': 'white',
    'axes.facecolor': '#FAFAFA',
    'axes.grid': True,
    'grid.alpha': 0.3,
    'grid.color': '#CCCCCC',
})

CHINESE_OS = '#E63946'
WESTERN_PROP = '#457B9D'
WESTERN_OS = '#2A9D8F'

df = pd.read_csv('chinese_distillation_rl_data.csv')


# ── Chart 1: RL Post-Training Share by Category ──

def chart1_rl_share_by_category():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 7), gridspec_kw={'width_ratios': [1.2, 1]})

    cat_colors = {
        'Chinese Open-Source': CHINESE_OS,
        'Western Proprietary': WESTERN_PROP,
        'Western Open-Source': WESTERN_OS,
    }

    rl_data = df[['Model Name', 'Category', 'Post-Training RL Share (Est %)']].copy()
    rl_data = rl_data.dropna(subset=['Post-Training RL Share (Est %)'])
    rl_data = rl_data.sort_values('Post-Training RL Share (Est %)', ascending=True)

    colors = [cat_colors.get(c, '#999999') for c in rl_data['Category']]
    bars = ax1.barh(range(len(rl_data)), rl_data['Post-Training RL Share (Est %)'],
                    color=colors, edgecolor='white', linewidth=0.5, height=0.7)

    ax1.set_yticks(range(len(rl_data)))
    ax1.set_yticklabels(rl_data['Model Name'], fontsize=9)
    ax1.set_xlabel('Estimated RL Share of Post-Training Effort (%)')
    ax1.set_title('RL Intensity in Post-Training by Model')
    ax1.set_xlim(0, 105)

    for i, (val, name) in enumerate(zip(rl_data['Post-Training RL Share (Est %)'], rl_data['Model Name'])):
        ax1.text(val + 1, i, f'{int(val)}%', va='center', fontsize=8, fontweight='bold')

    legend_patches = [mpatches.Patch(color=v, label=k) for k, v in cat_colors.items()]
    ax1.legend(handles=legend_patches, loc='lower right', fontsize=9, framealpha=0.9)

    cat_avg = df.groupby('Category')['Post-Training RL Share (Est %)'].agg(['mean', 'std', 'count']).reset_index()
    cat_avg = cat_avg.sort_values('mean', ascending=False)
    cat_colors_list = [cat_colors.get(c, '#999999') for c in cat_avg['Category']]

    bars2 = ax2.bar(range(len(cat_avg)), cat_avg['mean'], color=cat_colors_list,
                    edgecolor='white', linewidth=1.5, width=0.6)
    ax2.errorbar(range(len(cat_avg)), cat_avg['mean'], yerr=cat_avg['std'],
                 fmt='none', color='#333333', capsize=5, capthick=1.5)

    ax2.set_xticks(range(len(cat_avg)))
    ax2.set_xticklabels(cat_avg['Category'], fontsize=9, rotation=15, ha='right')
    ax2.set_ylabel('Average RL Share (%)')
    ax2.set_title('Average RL Intensity by Category')
    ax2.set_ylim(0, 100)

    for i, (m, s, n) in enumerate(zip(cat_avg['mean'], cat_avg['std'], cat_avg['count'])):
        ax2.text(i, m + s + 2, f'{m:.0f}%\n(n={int(n)})', ha='center', fontsize=10, fontweight='bold')

    plt.suptitle('Reinforcement Learning Focus: Chinese Open-Source vs. Western Models',
                 fontsize=16, fontweight='bold', y=1.02)
    plt.tight_layout()
    plt.savefig('01_rl_intensity_by_model_and_category.png', dpi=200, bbox_inches='tight')
    plt.close()
    print('Saved: 01_rl_intensity_by_model_and_category.png')


# ── Chart 2: RL Method Evolution Timeline ──

def chart2_rl_timeline():
    fig, ax = plt.subplots(figsize=(18, 9))

    timeline = [
        ('2023-03', 'GPT-4: RLHF (PPO) at scale', 'Western Proprietary', 'Established RLHF\nas standard'),
        ('2023-05', 'DPO Published', 'Western Open-Source', 'Simplified preference\noptimization'),
        ('2023-09', 'Baichuan2: Early Chinese RLHF', 'Chinese Open-Source', 'First Chinese\nopen RLHF'),
        ('2024-06', 'Constitutional AI scaled', 'Western Proprietary', 'AI-generated\nfeedback'),
        ('2024-09', 'Qwen2.5: PPO + DPO', 'Chinese Open-Source', 'Multi-method\nalignment'),
        ('2024-12', 'DeepSeek-V3: R1 distillation', 'Chinese Open-Source', 'Cross-model\nRL distillation'),
        ('2025-01', 'DeepSeek-R1-Zero:\nPure RL reasoning', 'Chinese Open-Source', 'PARADIGM SHIFT\nReasoning from RL alone'),
        ('2025-03', 'QwQ-32B: RL reasoning\nat 32B scale', 'Chinese Open-Source', 'Small-scale\nRL reasoning'),
        ('2025-05', 'Qwen3: Dual-mode\nRL reasoning', 'Chinese Open-Source', 'Think/chat\nmode switching'),
        ('2025-06', 'MiniMax-M1:\nCISPO algorithm', 'Chinese Open-Source', 'Novel RL\nbeats GRPO'),
        ('2025-07', 'Kimi-K2:\nAgentic RL', 'Chinese Open-Source', 'RL for\ntool use'),
        ('2025-12', 'OLMo 3: RLVR', 'Western Open-Source', 'Verifiable\nrewards'),
    ]

    cat_colors = {
        'Chinese Open-Source': CHINESE_OS,
        'Western Proprietary': WESTERN_PROP,
        'Western Open-Source': WESTERN_OS,
    }

    dates = pd.to_datetime([t[0] for t in timeline])
    x_positions = np.linspace(0.5, 11.5, len(timeline))

    ax.axhline(y=0, color='#333333', linewidth=2, zorder=1)

    for i, (date_str, label, category, desc) in enumerate(timeline):
        x = x_positions[i]
        color = cat_colors[category]

        direction = 1 if i % 2 == 0 else -1
        y_text = direction * (1.8 + (i % 3) * 0.6)

        ax.plot(x, 0, 'o', color=color, markersize=12, zorder=5, markeredgecolor='white', markeredgewidth=2)
        ax.plot([x, x], [0, y_text * 0.6], color=color, linewidth=1.5, alpha=0.6, zorder=2)

        fontweight = 'bold' if 'PARADIGM' in desc or 'Pure RL' in label else 'normal'
        fontsize = 8.5 if 'PARADIGM' in desc else 7.5

        bbox_props = dict(boxstyle='round,pad=0.4', facecolor=color, alpha=0.12, edgecolor=color, linewidth=1)
        ax.text(x, y_text, f'{label}\n{date_str}',
                ha='center', va='center', fontsize=fontsize, fontweight=fontweight,
                bbox=bbox_props, color='#1a1a1a')

    legend_patches = [mpatches.Patch(color=v, label=k) for k, v in cat_colors.items()]
    ax.legend(handles=legend_patches, loc='upper left', fontsize=10, framealpha=0.9)

    ax.set_xlim(-0.5, 12.5)
    ax.set_ylim(-5, 5)
    ax.set_xticks([])
    ax.set_yticks([])
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_visible(False)
    ax.set_facecolor('white')

    ax.set_title('RL Method Innovation Timeline: Chinese Labs Lead Reasoning-Focused RL',
                 fontsize=15, fontweight='bold', pad=20)

    plt.tight_layout()
    plt.savefig('02_rl_innovation_timeline.png', dpi=200, bbox_inches='tight')
    plt.close()
    print('Saved: 02_rl_innovation_timeline.png')


# ── Chart 3: RL Goal Comparison (Create vs Align) ──

def chart3_rl_goal_comparison():
    fig, ax = plt.subplots(figsize=(14, 8))

    goal_mapping = {
        'Create reasoning from scratch': 'Create Reasoning',
        'Create reasoning capability': 'Create Reasoning',
        'Create reasoning at 32B scale': 'Create Reasoning',
        'Create dual-mode reasoning': 'Create Reasoning',
        'Create tool-use and autonomous action capability': 'Create Agentic Capability',
        'Create reasoning across diverse tasks': 'Create Reasoning',
        'Infuse reasoning into general-purpose model': 'Transfer Reasoning (Distillation)',
        'Alignment + capability': 'Alignment + Capability',
        'Alignment and tool-calling': 'Alignment + Capability',
        'Alignment + tool-use': 'Alignment + Capability',
        'Alignment': 'Alignment Only',
        'Align behavior and reduce harm': 'Alignment + Safety',
        'Align behavior + safety + self-critique': 'Alignment + Safety',
        'Align behavior': 'Alignment + Safety',
        'Alignment + instruction following': 'Alignment Only',
        'Alignment + reasoning for small model': 'Alignment + Capability',
        'Transfer proprietary capabilities to open model': 'Transfer Reasoning (Distillation)',
        'Reasoning via verifiable rewards': 'Create Reasoning',
    }

    df['RL Goal Mapped'] = df['RL Goal'].map(goal_mapping).fillna('Other')

    goal_cat = df.groupby(['Category', 'RL Goal Mapped']).size().reset_index(name='count')

    goal_order = ['Create Reasoning', 'Create Agentic Capability', 'Transfer Reasoning (Distillation)',
                  'Alignment + Capability', 'Alignment + Safety', 'Alignment Only']
    goal_colors = {
        'Create Reasoning': '#E63946',
        'Create Agentic Capability': '#F4845F',
        'Transfer Reasoning (Distillation)': '#F7B267',
        'Alignment + Capability': '#7EC8E3',
        'Alignment + Safety': '#457B9D',
        'Alignment Only': '#1D3557',
    }

    categories = ['Chinese Open-Source', 'Western Proprietary', 'Western Open-Source']
    x = np.arange(len(categories))
    width = 0.12

    for j, goal in enumerate(goal_order):
        vals = []
        for cat in categories:
            match = goal_cat[(goal_cat['Category'] == cat) & (goal_cat['RL Goal Mapped'] == goal)]
            vals.append(match['count'].sum() if len(match) > 0 else 0)
        offset = (j - len(goal_order) / 2 + 0.5) * width
        bars = ax.bar(x + offset, vals, width, label=goal, color=goal_colors.get(goal, '#999'),
                      edgecolor='white', linewidth=0.5)
        for bar, val in zip(bars, vals):
            if val > 0:
                ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.05,
                        str(int(val)), ha='center', va='bottom', fontsize=8, fontweight='bold')

    ax.set_xticks(x)
    ax.set_xticklabels(categories, fontsize=11)
    ax.set_ylabel('Number of Models')
    ax.set_title('RL Goal Distribution: What RL Is Used For\n'
                 'Chinese labs use RL to CREATE capabilities; Western labs use RL to ALIGN behavior',
                 fontsize=13, fontweight='bold')
    ax.legend(title='RL Goal', fontsize=9, title_fontsize=10, loc='upper right', ncol=2)
    ax.set_ylim(0, ax.get_ylim()[1] * 1.15)

    plt.tight_layout()
    plt.savefig('03_rl_goal_distribution.png', dpi=200, bbox_inches='tight')
    plt.close()
    print('Saved: 03_rl_goal_distribution.png')


# ── Chart 4: Distillation Strategy Comparison ──

def chart4_distillation_strategy():
    fig, ax = plt.subplots(figsize=(15, 9))

    distill_data = [
        {'model': 'DeepSeek R1', 'category': 'Chinese Open-Source', 'source_size': 671,
         'targets': [1.5, 7, 8, 14, 32, 70], 'type': 'Cross-lab\n(into Qwen/Llama bases)',
         'color': CHINESE_OS, 'samples': '800K RL samples'},
        {'model': 'Qwen3', 'category': 'Chinese Open-Source', 'source_size': 235,
         'targets': [0.6, 1.7, 4, 8, 14, 32], 'type': 'Self-cascade\n(same family)',
         'color': CHINESE_OS, 'samples': 'Internal'},
        {'model': 'Qwen2.5', 'category': 'Chinese Open-Source', 'source_size': 72,
         'targets': [0.5, 1.5, 3, 7, 14, 32], 'type': 'Self-cascade\n(same family)',
         'color': CHINESE_OS, 'samples': 'Internal'},
        {'model': 'Gemma 3\n(from Gemini)', 'category': 'Western Open-Source', 'source_size': 500,
         'targets': [0.27, 1, 4, 12, 27], 'type': 'Proprietary→Open\n(knowledge distillation)',
         'color': WESTERN_OS, 'samples': 'Internal'},
        {'model': 'Llama 3.1', 'category': 'Western Open-Source', 'source_size': 405,
         'targets': [8, 70], 'type': 'Synthetic data\n(not traditional distillation)',
         'color': WESTERN_OS, 'samples': 'Synthetic SFT'},
        {'model': 'GPT-4 → mini', 'category': 'Western Proprietary', 'source_size': 1800,
         'targets': [200], 'type': 'Internal product tier\n(API-gated)',
         'color': WESTERN_PROP, 'samples': 'Proprietary'},
        {'model': 'Claude → Haiku', 'category': 'Western Proprietary', 'source_size': 800,
         'targets': [100], 'type': 'Internal product tier\n(API-gated)',
         'color': WESTERN_PROP, 'samples': 'Proprietary'},
    ]

    y_positions = list(range(len(distill_data)))[::-1]

    for i, d in enumerate(distill_data):
        y = y_positions[i]

        ax.plot(np.log10(d['source_size']), y, 's', color=d['color'],
                markersize=16, markeredgecolor='white', markeredgewidth=2, zorder=5)
        ax.text(np.log10(d['source_size']), y + 0.35, f"{d['source_size']}B",
                ha='center', va='bottom', fontsize=8, fontweight='bold', color=d['color'])

        for t in d['targets']:
            ax.plot(np.log10(t), y, 'o', color=d['color'],
                    markersize=8, markeredgecolor='white', markeredgewidth=1, zorder=5, alpha=0.7)
            ax.annotate('', xy=(np.log10(t), y), xytext=(np.log10(d['source_size']), y),
                        arrowprops=dict(arrowstyle='->', color=d['color'], alpha=0.4, lw=1.5))

        ax.text(-0.3, y, d['model'], ha='right', va='center', fontsize=10, fontweight='bold')
        ax.text(np.log10(max(d['targets'])) + 0.15, y - 0.15, d['type'],
                ha='left', va='center', fontsize=7.5, color='#555555', style='italic')

    ax.set_yticks([])
    xtick_vals = [0.5, 1, 3, 7, 14, 32, 70, 200, 405, 671, 1800]
    ax.set_xticks([np.log10(v) for v in xtick_vals])
    ax.set_xticklabels([f'{v}B' for v in xtick_vals], fontsize=9)
    ax.set_xlabel('Model Size (Parameters, log scale)', fontsize=12)
    ax.set_title('Distillation Strategies: Source → Target Model Sizes\n'
                 'Chinese labs distill broadly into open models; Western labs distill internally',
                 fontsize=13, fontweight='bold')
    ax.set_xlim(-0.6, 3.5)

    legend_patches = [
        mpatches.Patch(color=CHINESE_OS, label='Chinese Open-Source'),
        mpatches.Patch(color=WESTERN_OS, label='Western Open-Source'),
        mpatches.Patch(color=WESTERN_PROP, label='Western Proprietary'),
    ]
    ax.legend(handles=legend_patches, loc='lower right', fontsize=10)

    plt.tight_layout()
    plt.savefig('04_distillation_strategy_comparison.png', dpi=200, bbox_inches='tight')
    plt.close()
    print('Saved: 04_distillation_strategy_comparison.png')


# ── Chart 5: Reasoning Benchmark Performance ──

def chart5_benchmark_performance():
    fig, axes = plt.subplots(1, 3, figsize=(18, 7))

    bench_data = [
        ('DeepSeek-R1', 'Chinese Open-Source', 79.8, 97.3, 90.8),
        ('MiniMax-M1', 'Chinese Open-Source', 86.0, 96.8, 81.1),
        ('Kimi-K2', 'Chinese Open-Source', 69.6, 97.4, 89.5),
        ('Qwen3-235B', 'Chinese Open-Source', None, 96.2, 87.0),
        ('QwQ-32B', 'Chinese Open-Source', 44.0, 90, 85),
        ('R1-Distill-32B', 'Chinese Open-Source', 72.6, 94.3, 85.1),
        ('GPT-4o', 'Western Proprietary', 60, 90, 88),
        ('Llama 3.1 405B', 'Western Open-Source', 33, 73.8, 88.6),
        ('Phi-4 14B', 'Western Open-Source', None, 80.4, 84.8),
        ('Gemma 3 27B', 'Western Open-Source', None, None, 75),
    ]

    bench_df = pd.DataFrame(bench_data, columns=['Model', 'Category', 'AIME_2024', 'MATH_500', 'MMLU'])

    cat_colors = {
        'Chinese Open-Source': CHINESE_OS,
        'Western Proprietary': WESTERN_PROP,
        'Western Open-Source': WESTERN_OS,
    }

    for ax_idx, (col, title) in enumerate([('AIME_2024', 'AIME 2024 (%)'),
                                            ('MATH_500', 'MATH-500'),
                                            ('MMLU', 'MMLU')]):
        ax = axes[ax_idx]
        subset = bench_df.dropna(subset=[col]).sort_values(col, ascending=True)
        colors = [cat_colors.get(c, '#999') for c in subset['Category']]

        bars = ax.barh(range(len(subset)), subset[col], color=colors,
                       edgecolor='white', linewidth=0.5, height=0.6)

        ax.set_yticks(range(len(subset)))
        ax.set_yticklabels(subset['Model'], fontsize=9)
        ax.set_xlabel(title)
        ax.set_title(title, fontsize=12, fontweight='bold')

        for i, val in enumerate(subset[col]):
            ax.text(val + 0.3, i, f'{val:.1f}', va='center', fontsize=8)

    legend_patches = [mpatches.Patch(color=v, label=k) for k, v in cat_colors.items()]
    axes[2].legend(handles=legend_patches, loc='lower right', fontsize=8)

    plt.suptitle('Reasoning Benchmark Performance: RL-Trained Chinese Models Dominate',
                 fontsize=15, fontweight='bold', y=1.02)
    plt.tight_layout()
    plt.savefig('05_benchmark_performance_comparison.png', dpi=200, bbox_inches='tight')
    plt.close()
    print('Saved: 05_benchmark_performance_comparison.png')


# ── Chart 6: RL Method Ecosystem Map ──

def chart6_rl_method_map():
    fig, ax = plt.subplots(figsize=(16, 10))

    methods = [
        {'name': 'RLHF\n(PPO)', 'x': 2, 'y': 8, 'size': 2000, 'color': WESTERN_PROP,
         'origin': 'OpenAI (2022)', 'goal': 'Alignment'},
        {'name': 'Constitutional\nAI (RLAIF)', 'x': 4, 'y': 8.5, 'size': 1500, 'color': WESTERN_PROP,
         'origin': 'Anthropic (2023)', 'goal': 'Safety'},
        {'name': 'DPO', 'x': 6, 'y': 7, 'size': 2500, 'color': WESTERN_OS,
         'origin': 'Stanford (2023)', 'goal': 'Alignment'},
        {'name': 'Rejection\nSampling', 'x': 3, 'y': 5.5, 'size': 1200, 'color': WESTERN_OS,
         'origin': 'Meta (2024)', 'goal': 'Alignment'},
        {'name': 'RLVR', 'x': 8, 'y': 6.5, 'size': 1000, 'color': WESTERN_OS,
         'origin': 'AI2 (2025)', 'goal': 'Reasoning'},
        {'name': 'GRPO', 'x': 5, 'y': 3, 'size': 2800, 'color': CHINESE_OS,
         'origin': 'DeepSeek (2025)', 'goal': 'Reasoning'},
        {'name': 'Pure RL\n(no SFT)', 'x': 3, 'y': 2, 'size': 2200, 'color': CHINESE_OS,
         'origin': 'DeepSeek (2025)', 'goal': 'Reasoning'},
        {'name': 'CISPO', 'x': 7, 'y': 2.5, 'size': 1800, 'color': CHINESE_OS,
         'origin': 'MiniMax (2025)', 'goal': 'Reasoning'},
        {'name': 'Agentic\nRL', 'x': 9, 'y': 3.5, 'size': 1600, 'color': CHINESE_OS,
         'origin': 'Moonshot (2025)', 'goal': 'Agentic'},
        {'name': 'Dual-mode\nRL', 'x': 7, 'y': 4.5, 'size': 1400, 'color': CHINESE_OS,
         'origin': 'Alibaba (2025)', 'goal': 'Reasoning'},
    ]

    for m in methods:
        ax.scatter(m['x'], m['y'], s=m['size'], c=m['color'], alpha=0.25, zorder=2)
        ax.scatter(m['x'], m['y'], s=100, c=m['color'], alpha=0.9, zorder=3,
                   edgecolors='white', linewidth=1.5)
        ax.text(m['x'], m['y'] + 0.55, m['name'], ha='center', va='bottom',
                fontsize=9, fontweight='bold', color='#1a1a1a')
        ax.text(m['x'], m['y'] - 0.45, m['origin'], ha='center', va='top',
                fontsize=7, color='#666666', style='italic')

    ax.annotate('', xy=(5, 3.5), xytext=(2, 7.5),
                arrowprops=dict(arrowstyle='->', color='#999', lw=2, ls='--'))
    ax.text(3.2, 5.8, 'Evolution:\nAlignment → Reasoning', fontsize=8,
            color='#777', rotation=-50, ha='center')

    ax.axhspan(5, 10, alpha=0.04, color=WESTERN_PROP, zorder=0)
    ax.axhspan(1, 5, alpha=0.04, color=CHINESE_OS, zorder=0)
    ax.text(0.5, 9, 'ALIGNMENT\nFOCUS', fontsize=11, fontweight='bold',
            color=WESTERN_PROP, alpha=0.5, va='top')
    ax.text(0.5, 1.5, 'REASONING\nFOCUS', fontsize=11, fontweight='bold',
            color=CHINESE_OS, alpha=0.5, va='bottom')

    ax.set_xlim(0, 10.5)
    ax.set_ylim(1, 10)
    ax.set_xticks([])
    ax.set_yticks([])
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_visible(False)

    legend_patches = [
        mpatches.Patch(color=CHINESE_OS, label='Chinese Open-Source'),
        mpatches.Patch(color=WESTERN_PROP, label='Western Proprietary'),
        mpatches.Patch(color=WESTERN_OS, label='Western Open-Source'),
    ]
    ax.legend(handles=legend_patches, loc='upper right', fontsize=10, framealpha=0.9)

    ax.set_title('RL Method Landscape: Alignment vs. Reasoning\n'
                 'Bubble size ∝ estimated impact; position reflects primary goal',
                 fontsize=14, fontweight='bold')

    plt.tight_layout()
    plt.savefig('06_rl_method_landscape.png', dpi=200, bbox_inches='tight')
    plt.close()
    print('Saved: 06_rl_method_landscape.png')


# ── Chart 7: Distillation Flow Sankey-style ──

def chart7_distillation_heatmap():
    fig, ax = plt.subplots(figsize=(14, 8))

    sizes = ['<1B', '1-3B', '4-8B', '14-16B', '27-34B', '70-72B', '200B+']
    categories = ['Chinese Open-Source', 'Western Open-Source', 'Western Proprietary']

    availability = np.array([
        [3, 4, 5, 4, 4, 3, 5],
        [2, 2, 3, 2, 2, 1, 3],
        [0, 0, 0, 0, 0, 0, 3],
    ])

    cmap = sns.color_palette('YlOrRd', as_cmap=True)
    im = ax.imshow(availability, cmap=cmap, aspect='auto', vmin=0, vmax=6)

    ax.set_xticks(range(len(sizes)))
    ax.set_xticklabels(sizes, fontsize=11)
    ax.set_yticks(range(len(categories)))
    ax.set_yticklabels(categories, fontsize=11)
    ax.set_xlabel('Model Size Range', fontsize=12)

    for i in range(len(categories)):
        for j in range(len(sizes)):
            val = availability[i, j]
            text_color = 'white' if val > 3 else 'black'
            labels = {0: 'None', 1: 'Limited', 2: 'Some', 3: 'Good', 4: 'Strong', 5: 'Extensive'}
            ax.text(j, i, labels.get(val, str(val)), ha='center', va='center',
                    fontsize=10, fontweight='bold', color=text_color)

    ax.set_title('Open Model Availability by Size & Category\n'
                 'Chinese labs provide RL-enhanced models across all size tiers',
                 fontsize=13, fontweight='bold')

    cbar = plt.colorbar(im, ax=ax, shrink=0.6, pad=0.02)
    cbar.set_label('Availability Level', fontsize=10)

    plt.tight_layout()
    plt.savefig('07_model_availability_heatmap.png', dpi=200, bbox_inches='tight')
    plt.close()
    print('Saved: 07_model_availability_heatmap.png')


# ── Chart 8: Summary Radar Chart ──

def chart8_summary_radar():
    fig, ax = plt.subplots(figsize=(10, 10), subplot_kw=dict(polar=True))

    categories_radar = ['RL for\nReasoning', 'RL for\nAlignment', 'Distillation\nBreadth',
                        'Open\nAvailability', 'Benchmark\nPerformance', 'Compute\nEfficiency']
    N = len(categories_radar)
    angles = [n / float(N) * 2 * np.pi for n in range(N)]
    angles += angles[:1]

    chinese_os = [9, 4, 9, 10, 9, 9]
    western_prop = [5, 9, 3, 1, 7, 3]
    western_os = [5, 6, 5, 8, 6, 6]

    chinese_os += chinese_os[:1]
    western_prop += western_prop[:1]
    western_os += western_os[:1]

    ax.plot(angles, chinese_os, 'o-', color=CHINESE_OS, linewidth=2.5, label='Chinese Open-Source')
    ax.fill(angles, chinese_os, alpha=0.15, color=CHINESE_OS)

    ax.plot(angles, western_prop, 's-', color=WESTERN_PROP, linewidth=2.5, label='Western Proprietary')
    ax.fill(angles, western_prop, alpha=0.15, color=WESTERN_PROP)

    ax.plot(angles, western_os, '^-', color=WESTERN_OS, linewidth=2.5, label='Western Open-Source')
    ax.fill(angles, western_os, alpha=0.15, color=WESTERN_OS)

    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(categories_radar, fontsize=11)
    ax.set_ylim(0, 10)
    ax.set_yticks([2, 4, 6, 8, 10])
    ax.set_yticklabels(['2', '4', '6', '8', '10'], fontsize=8)

    ax.legend(loc='upper right', bbox_to_anchor=(1.3, 1.1), fontsize=11)
    ax.set_title('Comparative Strengths: RL & Distillation\nAcross LLM Ecosystems',
                 fontsize=14, fontweight='bold', pad=30)

    plt.tight_layout()
    plt.savefig('08_comparative_radar.png', dpi=200, bbox_inches='tight')
    plt.close()
    print('Saved: 08_comparative_radar.png')


if __name__ == '__main__':
    print('Generating visualizations...\n')
    chart1_rl_share_by_category()
    chart2_rl_timeline()
    chart3_rl_goal_comparison()
    chart4_distillation_strategy()
    chart5_benchmark_performance()
    chart6_rl_method_map()
    chart7_distillation_heatmap()
    chart8_summary_radar()
    print('\nAll charts generated successfully.')
