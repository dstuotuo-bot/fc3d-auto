"""
福彩3D 顶尖专业级分析系统（混合模式：规则推演 + AI验证）
功能：自动抓取 + 规则推演 + AI二次筛选 + 5注推荐 + 走势图
"""

import pandas as pd
import numpy as np
from collections import Counter
from datetime import datetime
import requests
from bs4 import BeautifulSoup
import os
import json
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import warnings
warnings.filterwarnings('ignore')

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'WenQuanYi Micro Hei', 'SimHei', 'Arial']
plt.rcParams['axes.unicode_minus'] = False

print("=" * 80)
print("福彩3D 顶尖专业级分析系统（混合模式：规则推演 + AI验证）")
print(f"运行时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
print("=" * 80)


# ============================================================================
# 抓取数据
# ============================================================================
def fetch_data():
    """抓取福彩3D历史数据"""
    all_data = []
    headers = {'User-Agent': 'Mozilla/5.0'}
    for page in range(1, 10):
        url = f"http://kaijiang.zhcw.com/zhcw/html/3d/list_{page}.html"
        try:
            r = requests.get(url, headers=headers, timeout=10)
            soup = BeautifulSoup(r.text, 'html.parser')
            rows = soup.find_all('tr')
            for row in rows[2:]:
                cols = row.find_all('td')
                if len(cols) >= 3:
                    period = cols[1].text.strip()
                    date = cols[0].text.strip()
                    em = row.find_all('em')
                    if len(em) >= 3:
                        nums = f"{em[0].text.strip()}{em[1].text.strip()}{em[2].text.strip()}"
                        all_data.append({'期号': period, '开奖日期': date, '开奖号码': nums})
        except:
            pass
    seen = set()
    unique = []
    for d in all_data:
        if d['期号'] not in seen:
            seen.add(d['期号'])
            unique.append(d)
    unique.sort(key=lambda x: int(x['期号']), reverse=True)
    return unique


def update_data():
    file_path = 'fc3d_data.xlsx'
    print("\n正在抓取最新数据...")
    data = fetch_data()
    if not data:
        print("抓取失败")
        return False
    df = pd.DataFrame(data)
    df.to_excel(file_path, index=False)
    print(f"成功抓取 {len(data)} 期数据")
    return True


def load_data():
    file_path = 'fc3d_data.xlsx'
    if not os.path.exists(file_path):
        if not update_data():
            return None
    df = pd.read_excel(file_path, dtype={'开奖号码': str})
    df['开奖号码'] = df['开奖号码'].str.replace(' ', '')
    df['开奖日期'] = pd.to_datetime(df['开奖日期'])
    df = df.sort_values('开奖日期').reset_index(drop=True)
    df['百位'] = df['开奖号码'].str[0].astype(int)
    df['十位'] = df['开奖号码'].str[1].astype(int)
    df['个位'] = df['开奖号码'].str[2].astype(int)
    df['和值'] = df['百位'] + df['十位'] + df['个位']
    df['跨度'] = df.apply(lambda x: max(x['百位'], x['十位'], x['个位']) - min(x['百位'], x['十位'], x['个位']), axis=1)
    def get_pattern(row):
        nums = [row['百位'], row['十位'], row['个位']]
        unique = len(set(nums))
        if unique == 1:
            return "豹子"
        elif unique == 2:
            return "组三"
        else:
            return "组六"
    df['形态'] = df.apply(get_pattern, axis=1)
    return df


# ============================================================================
# AI 调用
# ============================================================================
def call_deepseek(prompt):
    """调用 DeepSeek API"""
    api_key = os.environ.get('DEEPSEEK_API_KEY')
    if not api_key:
        print("  ⚠️ 未设置 DEEPSEEK_API_KEY，跳过 AI 验证")
        return None
    
    try:
        response = requests.post(
            'https://api.deepseek.com/v1/chat/completions',
            headers={
                'Authorization': f'Bearer {api_key}',
                'Content-Type': 'application/json'
            },
            json={
                'model': 'deepseek-chat',
                'messages': [
                    {'role': 'system', 'content': '你是福彩3D推演专家，基于历史数据统计规律给出推演建议。只输出JSON格式。'},
                    {'role': 'user', 'content': prompt}
                ],
                'temperature': 0.3,
                'max_tokens': 800
            },
            timeout=30
        )
        if response.status_code == 200:
            result = response.json()
            return result['choices'][0]['message']['content']
        else:
            print(f"  ⚠️ AI调用失败: {response.status_code}")
            return None
    except Exception as e:
        print(f"  ⚠️ AI调用异常: {e}")
        return None


def ai_verify_predictions(candidates, df_recent, features):
    """AI 验证推演结果"""
    # 准备历史数据摘要
    recent_10 = df_recent.tail(10)[['期号', '开奖号码', '和值', '跨度', '形态']].to_dict('records')
    history_summary = f"最近10期开奖: {recent_10}"
    
    # 准备候选号码
    candidates_summary = []
    for c in candidates[:10]:
        candidates_summary.append({
            '号码': c['号码'],
            '规则得分': c['得分'],
            '和值': c['和值'],
            '形态': c['形态'],
            '奇偶': c['奇偶'],
            '大小': c['大小']
        })
    
    prompt = f"""
请分析以下福彩3D数据，对候选号码进行验证和排序：

【历史数据】
{history_summary}

【常见特征】
最常见和值: {features['common_sum']}
最常见跨度: {features['span_cnt'].most_common(1)[0][0]}
最常见奇偶: {features['common_parity']}
最常见大小: {features['common_size']}
预测形态: {features.get('predicted_pattern', '组六')}

【候选号码】（已按规则得分排序）
{json.dumps(candidates_summary, ensure_ascii=False, indent=2)}

请输出JSON格式：
{{
    "verified_predictions": [
        {{"number": "号码", "ai_score": 0-100, "reason": "推荐理由"}}
    ],
    "analysis": "简要分析"
}}
"""
    
    result = call_deepseek(prompt)
    if result:
        try:
            # 提取JSON
            result = result.strip()
            if '```json' in result:
                result = result.split('```json')[1].split('```')[0]
            elif '```' in result:
                result = result.split('```')[1].split('```')[0]
            ai_data = json.loads(result)
            return ai_data
        except:
            print(f"  ⚠️ AI响应解析失败")
            return None
    return None


# ============================================================================
# 走势图生成
# ============================================================================
def generate_charts(df):
    """生成6张走势图"""
    print("\n" + "=" * 80)
    print("【生成走势图】")
    print("=" * 80)
    
    charts = []
    if len(df) < 15:
        print("数据不足15期，跳过走势图")
        return charts
    
    recent_15 = df.tail(15).copy()
    recent_15 = recent_15.reset_index(drop=True)
    
    # 图1：和值走势图
    try:
        fig, ax = plt.subplots(figsize=(16, 8))
        sums = recent_15['和值'].tolist()
        periods = recent_15['期号'].tolist()
        ax.plot(range(len(sums)), sums, 'r-o', linewidth=2, markersize=8, color='#e74c3c')
        mean_sum = np.mean(sums)
        ax.axhline(y=mean_sum, color='#3498db', linestyle='--', linewidth=2, label=f'均值: {mean_sum:.1f}')
        ax.axhspan(9, 16, alpha=0.15, color='#2ecc71', label='常见区间 9-16')
        ax.set_title('和值走势图（最近15期）', fontsize=16)
        ax.set_xlabel('期数', fontsize=12)
        ax.set_ylabel('和值', fontsize=12)
        ax.set_xticks(range(0, len(sums), 3))
        ax.set_xticklabels([f"{periods[i]}" for i in range(0, len(sums), 3)], fontsize=10)
        ax.legend()
        ax.grid(True, alpha=0.3)
        plt.tight_layout()
        plt.savefig('fc3d_sum_trend.png', dpi=150, bbox_inches='tight')
        plt.close()
        charts.append('fc3d_sum_trend.png')
        print("  ✅ 和值走势图")
    except Exception as e:
        print(f"  ⚠️ 和值走势图失败: {e}")
    
    # 图2：跨度走势图
    try:
        fig, ax = plt.subplots(figsize=(16, 8))
        spans = recent_15['跨度'].tolist()
        periods = recent_15['期号'].tolist()
        ax.plot(range(len(spans)), spans, 'g-s', linewidth=2, markersize=8, color='#27ae60')
        mean_span = np.mean(spans)
        ax.axhline(y=mean_span, color='#e67e22', linestyle='--', linewidth=2, label=f'均值: {mean_span:.1f}')
        ax.set_title('跨度走势图（最近15期）', fontsize=16)
        ax.set_xlabel('期数', fontsize=12)
        ax.set_ylabel('跨度', fontsize=12)
        ax.set_xticks(range(0, len(spans), 3))
        ax.set_xticklabels([f"{periods[i]}" for i in range(0, len(spans), 3)], fontsize=10)
        ax.set_ylim(-0.5, 9.5)
        ax.legend()
        ax.grid(True, alpha=0.3)
        plt.tight_layout()
        plt.savefig('fc3d_span_trend.png', dpi=150, bbox_inches='tight')
        plt.close()
        charts.append('fc3d_span_trend.png')
        print("  ✅ 跨度走势图")
    except Exception as e:
        print(f"  ⚠️ 跨度走势图失败: {e}")
    
    # 图3：和值分布直方图
    try:
        fig, ax = plt.subplots(figsize=(20, 10))
        all_sums = df['和值'].tolist()
        ax.hist(all_sums, bins=range(0, 28), edgecolor='white', alpha=0.8, color='#e74c3c', rwidth=0.9)
        ax.axvline(x=np.mean(all_sums), color='#3498db', linestyle='--', linewidth=2, label=f'均值: {np.mean(all_sums):.1f}')
        ax.set_title('和值分布直方图', fontsize=16)
        ax.set_xlabel('和值', fontsize=12)
        ax.set_ylabel('出现次数', fontsize=12)
        ax.legend()
        plt.tight_layout()
        plt.savefig('fc3d_sum_hist.png', dpi=150, bbox_inches='tight')
        plt.close()
        charts.append('fc3d_sum_hist.png')
        print("  ✅ 和值分布图")
    except Exception as e:
        print(f"  ⚠️ 和值分布图失败: {e}")
    
    # 图4：形态占比饼图
    try:
        fig, ax = plt.subplots(figsize=(10, 8))
        pattern_cnt = Counter(df['形态'])
        colors = {'组六': '#3498db', '组三': '#e74c3c', '豹子': '#f39c12'}
        colors_list = [colors.get(k, '#95a5a6') for k in pattern_cnt.keys()]
        ax.pie(pattern_cnt.values(), labels=pattern_cnt.keys(), autopct='%1.1f%%', colors=colors_list, shadow=True, startangle=90, textprops={'fontsize': 14})
        ax.set_title('形态占比', fontsize=16)
        plt.tight_layout()
        plt.savefig('fc3d_pattern_pie.png', dpi=150, bbox_inches='tight')
        plt.close()
        charts.append('fc3d_pattern_pie.png')
        print("  ✅ 形态占比图")
    except Exception as e:
        print(f"  ⚠️ 形态饼图失败: {e}")
    
    # 图5：各位置数字频率
    try:
        fig, axes = plt.subplots(1, 3, figsize=(18, 6))
        positions = ['百位', '十位', '个位']
        colors_bar = ['#e74c3c', '#3498db', '#27ae60']
        for i, pos in enumerate(positions):
            counts = df[pos].value_counts().sort_index()
            axes[i].bar(counts.index, counts.values, color=colors_bar[i], alpha=0.8, edgecolor='white')
            axes[i].set_title(f'{pos}位数字频率', fontsize=14)
            axes[i].set_xlabel('数字', fontsize=12)
            axes[i].set_ylabel('出现次数', fontsize=12)
            axes[i].set_xticks(range(10))
            axes[i].grid(True, alpha=0.2, axis='y')
            for bar, v in zip(axes[i].patches, counts.values):
                axes[i].text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5, str(v), ha='center', fontsize=11)
        fig.suptitle('各位置数字频率分布', fontsize=16)
        plt.tight_layout()
        plt.savefig('fc3d_position_freq.png', dpi=150, bbox_inches='tight')
        plt.close()
        charts.append('fc3d_position_freq.png')
        print("  ✅ 位置频率图")
    except Exception as e:
        print(f"  ⚠️ 位置频率图失败: {e}")
    
    # 图6：奇偶走势图
    try:
        fig, ax = plt.subplots(figsize=(16, 8))
        periods = recent_15['期号'].tolist()
        odd_counts = [(row['百位'] % 2) + (row['十位'] % 2) + (row['个位'] % 2) for _, row in recent_15.iterrows()]
        colors = ['#e74c3c' if x >= 3 else '#3498db' if x <= 0 else '#f39c12' for x in odd_counts]
        ax.bar(range(len(odd_counts)), odd_counts, color=colors, alpha=0.8, edgecolor='white')
        ax.axhline(y=3, color='#e74c3c', linestyle='--', linewidth=2, label='全奇线 (3个奇数)')
        ax.axhline(y=0, color='#3498db', linestyle='--', linewidth=2, label='全偶线 (0个奇数)')
        ax.set_title('奇偶个数走势图（最近15期）', fontsize=16)
        ax.set_xlabel('期数', fontsize=12)
        ax.set_ylabel('奇数个数', fontsize=12)
        ax.set_xticks(range(0, len(odd_counts), 3))
        ax.set_xticklabels([f"{periods[i]}" for i in range(0, len(odd_counts), 3)], fontsize=10)
        ax.set_yticks(range(4))
        ax.set_ylim(-0.5, 3.5)
        ax.legend()
        ax.grid(True, alpha=0.3)
        plt.tight_layout()
        plt.savefig('fc3d_parity_trend.png', dpi=150, bbox_inches='tight')
        plt.close()
        charts.append('fc3d_parity_trend.png')
        print("  ✅ 奇偶走势图")
    except Exception as e:
        print(f"  ⚠️ 奇偶走势图失败: {e}")
    
    print(f"\n✅ 共生成 {len(charts)} 张走势图")
    return charts


# ============================================================================
# 分析函数
# ============================================================================
def analyze_pattern(df):
    patterns = df['形态'].tolist()
    pattern_cnt = Counter(patterns)
    last_pattern = patterns[-1]
    recent_100 = df.tail(100)['形态'].tolist()
    transitions = {}
    for i in range(len(recent_100)-1):
        key = (recent_100[i], recent_100[i+1])
        transitions[key] = transitions.get(key, 0) + 1
    next_patterns = {}
    for (prev, next_p), cnt in transitions.items():
        if prev == last_pattern:
            next_patterns[next_p] = next_patterns.get(next_p, 0) + cnt
    if next_patterns:
        predicted_pattern = max(next_patterns, key=next_patterns.get)
    else:
        predicted_pattern = pattern_cnt.most_common(1)[0][0]
    return predicted_pattern, pattern_cnt


def analyze_positions(df, pos_data):
    periods = len(df)
    for pos in ['百位', '十位', '个位']:
        total_counts = df[pos].value_counts().sort_index()
        miss = {}
        for n in range(10):
            idx = df[df[pos] == n].index.tolist()
            miss[n] = periods - idx[-1] - 1 if idx else periods
        recent30 = df.tail(30)[pos].tolist()
        recent_cnt = Counter(recent30)
        top3 = [n for n, _ in recent_cnt.most_common(3)]
        position_scores = {}
        for n in range(10):
            freq_score = total_counts.get(n, 0) / periods
            recent_score = recent_cnt.get(n, 0) / 30
            miss_score = 1 / (miss[n] + 1)
            total_score = freq_score * 0.4 + recent_score * 0.4 + miss_score * 0.2
            position_scores[n] = total_score
        pos_data[pos] = {'top3': top3, 'miss': miss, 'counts': total_counts, 'position_scores': position_scores}


def analyze_features(df):
    sums = df['和值'].tolist()
    sum_cnt = Counter(sums)
    spans = df['跨度'].tolist()
    span_cnt = Counter(spans)
    parity = []
    for _, r in df.iterrows():
        p = ('奇' if r['百位']%2 else '偶') + ('奇' if r['十位']%2 else '偶') + ('奇' if r['个位']%2 else '偶')
        parity.append(p)
    p_cnt = Counter(parity)
    size = []
    for _, r in df.iterrows():
        s = ('大' if r['百位']>=5 else '小') + ('大' if r['十位']>=5 else '小') + ('大' if r['个位']>=5 else '小')
        size.append(s)
    s_cnt = Counter(size)
    return {
        'common_sum': sum_cnt.most_common(1)[0][0],
        'common_sum_top5': [s for s, _ in sum_cnt.most_common(5)],
        'common_parity': p_cnt.most_common(1)[0][0],
        'common_size': s_cnt.most_common(1)[0][0],
        'span_cnt': span_cnt,
        'p_cnt': p_cnt,
        's_cnt': s_cnt
    }


def rule_predict(df, pos_data, features, predicted_pattern):
    """规则推演 - 生成27个候选并打分"""
    candidates = []
    for b in pos_data['百位']['top3']:
        for s in pos_data['十位']['top3']:
            for g in pos_data['个位']['top3']:
                num = f"{b}{s}{g}"
                total = b + s + g
                parity = ('奇' if b%2 else '偶') + ('奇' if s%2 else '偶') + ('奇' if g%2 else '偶')
                size = ('大' if b>=5 else '小') + ('大' if s>=5 else '小') + ('大' if g>=5 else '小')
                unique = len({b, s, g})
                if unique == 1:
                    pattern = "豹子"
                elif unique == 2:
                    pattern = "组三"
                else:
                    pattern = "组六"
                score = 0
                reasons = []
                if total == features['common_sum']:
                    score += 5
                    reasons.append("和值✓")
                elif total in features['common_sum_top5']:
                    score += 2
                    reasons.append("和值≈")
                if parity == features['common_parity']:
                    score += 4
                    reasons.append("奇偶✓")
                if size == features['common_size']:
                    score += 4
                    reasons.append("大小✓")
                if pattern == predicted_pattern:
                    score += 3
                    reasons.append("形态✓")
                pos_score = (pos_data['百位']['position_scores'].get(b, 0) +
                            pos_data['十位']['position_scores'].get(s, 0) +
                            pos_data['个位']['position_scores'].get(g, 0))
                pos_score_normalized = min(3, pos_score * 1.2)
                score += pos_score_normalized
                if pos_score_normalized > 2:
                    reasons.append("位置优")
                last_num = df.iloc[-1]['开奖号码']
                if num != last_num:
                    score += 1
                    reasons.append("防重✓")
                candidates.append({
                    '号码': num, '得分': round(score, 1), '和值': total,
                    '形态': pattern, '奇偶': parity, '大小': size, '原因': reasons
                })
    candidates.sort(key=lambda x: x['得分'], reverse=True)
    return candidates


def hybrid_predict(df, pos_data, features, predicted_pattern):
    """混合模式推演：规则推演 + AI验证"""
    
    # 第一步：规则推演
    print("\n" + "=" * 80)
    print("【第一步：规则推演】")
    print("=" * 80)
    candidates = rule_predict(df, pos_data, features, predicted_pattern)
    print(f"规则推演生成 {len(candidates)} 个候选号码")
    print(f"Top5 候选: {[c['号码'] for c in candidates[:5]]}")
    
    # 第二步：AI验证
    print("\n" + "=" * 80)
    print("【第二步：AI验证】")
    print("=" * 80)
    
    ai_result = ai_verify_predictions(candidates, df, features)
    
    # 第三步：综合排序
    print("\n" + "=" * 80)
    print("【第三步：综合排序】")
    print("=" * 80)
    
    if ai_result and 'verified_predictions' in ai_result:
        # 建立AI评分映射
        ai_scores = {}
        for item in ai_result['verified_predictions']:
            ai_scores[item['number']] = item.get('ai_score', 50)
        
        # 综合评分 = 规则分 * 0.6 + AI分 * 0.4
        for c in candidates:
            ai_score = ai_scores.get(c['号码'], 50)
            c['ai得分'] = ai_score
            c['综合得分'] = round(c['得分'] * 0.6 + ai_score * 0.4, 1)
            c['原因'].append(f"AI:{ai_score}分")
        
        # 按综合得分重新排序
        candidates.sort(key=lambda x: x['综合得分'], reverse=True)
        
        print("AI验证完成，已综合评分")
        if 'analysis' in ai_result:
            print(f"AI分析: {ai_result['analysis'][:100]}...")
    else:
        print("AI验证未生效，使用规则推演结果")
        for c in candidates:
            c['综合得分'] = c['得分']
    
    return candidates


# ============================================================================
# 生成HTML报告
# ============================================================================
def generate_html_report(df, pos_data, features, predicted_pattern, candidates, next_period, chart_files):
    recent_5 = df.tail(5)
    pos_scores = {}
    for pos in ['百位', '十位', '个位']:
        scores = pos_data[pos]['position_scores']
        sorted_scores = sorted(scores.items(), key=lambda x: x[1], reverse=True)[:5]
        pos_scores[pos] = sorted_scores
    group6_pct = features['p_cnt'].get('组六', 0) / len(df) * 100
    group3_pct = features['p_cnt'].get('组三', 0) / len(df) * 100
    group_bz_pct = features['p_cnt'].get('豹子', 0) / len(df) * 100
    
    charts_html = ''
    for chart in chart_files:
        charts_html += f'<div class="chart-card"><a href="{chart}" target="_blank"><img src="{chart}" alt="走势图"></a></div>'
    
    # 判断是否使用了AI
    use_ai = 'ai得分' in candidates[0] if candidates else False
    
    html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover">
    <title>福彩3D · 专业分析报告（AI增强版）</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: 'Microsoft YaHei', 'PingFang SC', Arial, sans-serif;
            background: linear-gradient(145deg, #0f172a 0%, #1e293b 100%);
            min-height: 100vh;
            padding: 20px;
        }}
        .container {{ max-width: 1400px; margin: 0 auto; }}
        .hero {{ text-align: center; margin-bottom: 30px; }}
        .hero h1 {{
            font-size: clamp(24px, 7vw, 42px);
            background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
            -webkit-background-clip: text;
            background-clip: text;
            color: transparent;
        }}
        .hero-badge {{
            display: flex;
            flex-wrap: wrap;
            justify-content: center;
            gap: 8px 16px;
            background: rgba(255,255,255,0.1);
            padding: 8px 20px;
            border-radius: 100px;
            font-size: 12px;
            color: #a5b4fc;
            margin-top: 10px;
        }}
        .ai-badge {{
            background: linear-gradient(135deg, #10b981, #059669);
            padding: 2px 8px;
            border-radius: 20px;
            font-size: 10px;
            margin-left: 10px;
        }}
        .tabs {{
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            margin-bottom: 20px;
            border-bottom: 1px solid rgba(255,255,255,0.1);
            padding-bottom: 10px;
        }}
        .tab-btn {{
            background: rgba(30,41,59,0.5);
            border: none;
            padding: 10px 20px;
            border-radius: 30px;
            color: #94a3b8;
            cursor: pointer;
            transition: all 0.3s;
            font-size: 14px;
        }}
        .tab-btn:hover {{ background: rgba(245,87,108,0.2); color: #f093fb; transform: translateY(-2px); }}
        .tab-btn.active {{ background: linear-gradient(135deg, #f093fb, #f5576c); color: white; }}
        .tab-content {{ display: none; animation: fadeIn 0.3s ease; }}
        .tab-content.active {{ display: block; }}
        @keyframes fadeIn {{ from {{ opacity: 0; }} to {{ opacity: 1; }} }}
        .glass-card {{
            background: rgba(30,41,59,0.7);
            backdrop-filter: blur(12px);
            border-radius: 24px;
            border: 1px solid rgba(255,255,255,0.1);
            margin-bottom: 20px;
            overflow: hidden;
            transition: all 0.3s ease;
        }}
        .glass-card:hover {{
            transform: translateY(-4px);
            box-shadow: 0 20px 40px rgba(0,0,0,0.3);
            border-color: rgba(245,87,108,0.3);
        }}
        .card-header {{ padding: 16px 20px; border-bottom: 1px solid rgba(255,255,255,0.08); display: flex; align-items: center; gap: 10px; }}
        .card-header h2 {{ font-size: 18px; color: #f1f5f9; }}
        .card-body {{ padding: 20px; }}
        .prediction-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 16px;
        }}
        @media (max-width: 640px) {{ .prediction-grid {{ grid-template-columns: 1fr; }} }}
        .prediction-item {{
            background: linear-gradient(145deg, #1e293b, #0f172a);
            border-radius: 28px;
            padding: 20px;
            text-align: center;
            border: 1px solid rgba(255,255,255,0.05);
            transition: all 0.3s ease;
            cursor: pointer;
        }}
        .prediction-item:hover {{ transform: scale(1.02); border-color: rgba(245,87,108,0.5); box-shadow: 0 0 20px rgba(245,87,108,0.2); }}
        .prediction-balls {{ display: flex; justify-content: center; gap: 12px; margin-bottom: 12px; }}
        .ball-large {{
            width: clamp(55px, 15vw, 70px);
            height: clamp(55px, 15vw, 70px);
            background: linear-gradient(145deg, #f093fb, #f5576c);
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: clamp(24px, 6vw, 32px);
            font-weight: 800;
            color: white;
            box-shadow: 0 8px 20px rgba(0,0,0,0.3);
        }}
        .prediction-score {{ display: inline-block; padding: 4px 14px; border-radius: 40px; font-size: 13px; font-weight: 600; }}
        .score-high {{ background: linear-gradient(135deg, #f5576c, #f093fb); color: white; }}
        .score-mid {{ background: linear-gradient(135deg, #f59e0b, #f97316); color: white; }}
        .score-low {{ background: linear-gradient(135deg, #10b981, #34d399); color: white; }}
        .ai-score {{ font-size: 10px; color: #10b981; margin-top: 4px; }}
        .stats-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
            gap: 16px;
            margin-bottom: 20px;
        }}
        @media (max-width: 640px) {{ .stats-grid {{ grid-template-columns: repeat(2, 1fr); }} }}
        .stat-card {{
            background: rgba(30,41,59,0.6);
            border-radius: 20px;
            padding: 16px;
            text-align: center;
            transition: all 0.3s ease;
            cursor: pointer;
        }}
        .stat-card:hover {{ transform: translateY(-2px); background: rgba(30,41,59,0.8); }}
        .stat-value {{ font-size: clamp(28px, 8vw, 36px); font-weight: 800; color: #f1f5f9; }}
        .stat-label {{ font-size: 12px; color: #94a3b8; }}
        .table-wrapper {{ overflow-x: auto; }}
        .data-table {{ width: 100%; border-collapse: collapse; min-width: 500px; }}
        .data-table th, .data-table td {{ padding: 12px 10px; color: #e2e8f0; border-bottom: 1px solid rgba(255,255,255,0.05); text-align: center; }}
        .data-table th {{ color: #94a3b8; font-weight: 500; }}
        .ball-small {{
            display: inline-flex; width: 34px; height: 34px;
            background: linear-gradient(145deg, #f093fb, #f5576c);
            border-radius: 50%; align-items: center; justify-content: center;
            font-weight: 700; font-size: 14px; color: white; margin: 0 2px;
        }}
        .progress-bar {{ background: rgba(255,255,255,0.1); border-radius: 12px; height: 24px; overflow: hidden; margin: 8px 0; }}
        .progress-fill {{
            background: linear-gradient(90deg, #f093fb, #f5576c);
            height: 100%; display: flex; align-items: center; padding-left: 8px;
            color: white; font-size: 11px; border-radius: 12px;
        }}
        .charts-grid {{ display: flex; flex-direction: column; gap: 24px; }}
        .chart-card {{ background: rgba(15,23,42,0.8); border-radius: 20px; padding: 16px; text-align: center; }}
        .chart-card img {{ width: 100%; border-radius: 12px; cursor: pointer; }}
        .footer {{ text-align: center; padding: 20px; color: #475569; font-size: 12px; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="hero">
            <h1>🎯 福彩3D · 专业分析报告<span class="ai-badge">AI增强版</span></h1>
            <div class="hero-badge">
                <span>📅 {datetime.now().strftime('%Y-%m-%d')}</span>
                <span>⚡ 下一期: {next_period}</span>
                <span>📊 基于 {len(df)} 期历史数据</span>
                <span>🤖 AI验证已启用</span>
            </div>
        </div>

        <!-- 推演结果 -->
        <div class="glass-card">
            <div class="card-header"><span>⭐</span><h2>智能推演 · 下一期预测（混合模式：规则+AI）</h2></div>
            <div class="card-body">
                <div class="prediction-grid">
'''
    
    for i, c in enumerate(candidates[:5], 1):
        nums = list(c['号码'])
        score_class = "score-high" if c['综合得分'] >= 16 else "score-mid" if c['综合得分'] >= 12 else "score-low"
        ai_badge = f'<div class="ai-score">AI:{c.get("ai得分", 50)}分</div>' if 'ai得分' in c else ''
        html += f'''
                    <div class="prediction-item">
                        <div class="prediction-balls">
                            <div class="ball-large">{nums[0]}</div>
                            <div class="ball-large">{nums[1]}</div>
                            <div class="ball-large">{nums[2]}</div>
                        </div>
                        <div><span class="prediction-score {score_class}">{c['综合得分']}分</span></div>
                        {ai_badge}
                        <div style="font-size: 11px; color:#94a3b8; margin-top: 5px;">形态:{c['形态']} | 和值:{c['和值']}</div>
                    </div>
'''
    
    html += f'''
                </div>
            </div>
        </div>

        <!-- 统计卡片 -->
        <div class="stats-grid">
            <div class="stat-card"><div class="stat-value">{features['common_sum']}</div><div class="stat-label">最常见和值</div></div>
            <div class="stat-card"><div class="stat-value">{features['span_cnt'].most_common(1)[0][0]}</div><div class="stat-label">最常见跨度</div></div>
            <div class="stat-card"><div class="stat-value">{predicted_pattern}</div><div class="stat-label">预测形态</div></div>
            <div class="stat-card"><div class="stat-value">AI</div><div class="stat-label">验证模式</div></div>
        </div>

        <!-- 最近5期 -->
        <div class="glass-card">
            <div class="card-header"><span>📋</span><h2>最近5期开奖记录</h2></div>
            <div class="card-body">
                <div class="table-wrapper">
                    <table class="data-table">
                        <thead><tr><th>期号</th><th>开奖日期</th><th>开奖号码</th><th>形态</th><th>和值</th><th>跨度</th></tr></thead>
                        <tbody>
'''
    
    for _, row in recent_5.iterrows():
        nums = row['开奖号码']
        html += f'<tr><td>{row["期号"]}</td><td>{row["开奖日期"].strftime("%Y-%m-%d")}</td><td><span class="ball-small">{nums[0]}</span><span class="ball-small">{nums[1]}</span><span class="ball-small">{nums[2]}</span></td><td>{row["形态"]}</td><td>{row["和值"]}</td><td>{row["跨度"]}</td></tr>'
    
    html += f'''
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- 走势图 -->
        <div class="glass-card">
            <div class="card-header"><span>📈</span><h2>走势图分析（点击图片可放大）</h2></div>
            <div class="card-body">
                <div class="charts-grid">
                    {charts_html}
                </div>
            </div>
        </div>

        <div class="footer">⚠️ 本分析基于历史数据统计规律 + AI辅助验证，仅供学习参考</div>
    </div>

    <script>
        document.querySelectorAll('.tab-btn').forEach(btn => {{
            btn.addEventListener('click', () => {{
                document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
                document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
                btn.classList.add('active');
                document.getElementById(btn.dataset.tab).classList.add('active');
            }});
        }});
    </script>
</body>
</html>
'''
    
    filename = "fc3d_report.html"
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(html)
    return filename


# ============================================================================
# 主程序
# ============================================================================
def main():
    df = load_data()
    if df is None:
        print("数据加载失败")
        return
    
    latest = df.iloc[-1]['期号']
    next_period = int(latest) + 1
    print(f"\n最新期号: {latest} | 下一期: {next_period}")
    
    df_recent = df.tail(100)
    print(f"分析基数: 最近{len(df_recent)}期")
    
    predicted_pattern, pattern_cnt = analyze_pattern(df_recent)
    print(f"预测形态: {predicted_pattern}")
    
    pos_data = {}
    analyze_positions(df_recent, pos_data)
    
    features = analyze_features(df_recent)
    features['predicted_pattern'] = predicted_pattern
    
    # 混合模式推演
    candidates = hybrid_predict(df_recent, pos_data, features, predicted_pattern)
    
    # 打印最终结果
    print("\n" + "=" * 80)
    print("【最终推荐】综合规则+AI评分")
    print("=" * 80)
    for i, c in enumerate(candidates[:5], 1):
        ai_info = f" | AI:{c.get('ai得分', '-')}分" if 'ai得分' in c else ""
        print(f"  第{i}注: {c['号码']} | 综合得分: {c['综合得分']}{ai_info} | 和值:{c['和值']}")
    
    chart_files = generate_charts(df)
    html_file = generate_html_report(df_recent, pos_data, features, predicted_pattern, candidates, next_period, chart_files)
    
    print(f"\n✅ HTML报告已生成: {html_file}")
    print(f"📁 位置: {os.path.abspath(html_file)}")


if __name__ == "__main__":
    main()
