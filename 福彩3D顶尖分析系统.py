"""
福彩3D 顶尖专业级分析系统（完整满血版 + 全端自适应）
功能：自动抓取 + 深度分析 + 5注推演 + HTML报告 + 历史准确率 + 走势图
"""

import pandas as pd
import numpy as np
from collections import Counter
from datetime import datetime
import requests
from bs4 import BeautifulSoup
import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import warnings
warnings.filterwarnings('ignore')

# 设置中文字体（兼容云端Linux环境）
plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'WenQuanYi Micro Hei', 'SimHei', 'Arial']
plt.rcParams['axes.unicode_minus'] = False

print("=" * 80)
print("福彩3D 顶尖专业级分析系统")
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
    """更新数据文件"""
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
    """加载数据"""
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
# 历史准确率统计
# ============================================================================
def calculate_accuracy(df):
    """计算历史推演准确率"""
    print("\n" + "=" * 80)
    print("【历史准确率统计】")
    print("=" * 80)
    
    total_periods = len(df)
    print(f"总期数: {total_periods}期")
    
    if total_periods < 30:
        print("⚠️ 数据不足（需要至少30期），继续积累数据")
        return None
    
    test_periods = min(30, total_periods // 3)
    if test_periods < 5:
        return None
    
    hit_count = 0
    
    for i in range(total_periods - test_periods, total_periods - 1):
        train_df = df.iloc[:i+1]
        actual = df.iloc[i+1]['开奖号码']
        
        if len(train_df) >= 20:
            recent_train = train_df.tail(50)
            top3 = {}
            for pos in ['百位', '十位', '个位']:
                recent30 = recent_train.tail(30)[pos].tolist()
                rc = Counter(recent30)
                top3[pos] = [n for n, _ in rc.most_common(3)]
            
            predictions = []
            for b in top3['百位']:
                for s in top3['十位']:
                    for g in top3['个位']:
                        predictions.append(f"{b}{s}{g}")
            
            if actual in predictions:
                hit_count += 1
    
    tested = total_periods - (total_periods - test_periods) - 1
    if tested <= 0:
        return None
    
    accuracy = hit_count / tested * 100
    
    print(f"\n测试期数: {tested}期")
    print(f"命中次数: {hit_count}期")
    print(f"准确率: {accuracy:.1f}%")
    
    return {'accuracy': accuracy, 'hit_count': hit_count, 'total_tested': tested}


# ============================================================================
# 走势图生成（完整版）
# ============================================================================
def generate_charts(df):
    """生成走势图 - 完整版"""
    print("\n" + "=" * 80)
    print("【生成走势图】")
    print("=" * 80)
    
    charts = []
    
    if len(df) < 15:
        print("  ⚠️ 数据不足15期，跳过走势图")
        return charts
    
    recent_15 = df.tail(15).copy()
    recent_15 = recent_15.reset_index(drop=True)
    
    # 图1：和值走势图
    try:
        fig, ax = plt.subplots(figsize=(14, 7))
        dates = [str(i+1) for i in range(len(recent_15))]
        sums = recent_15['和值'].tolist()
        
        ax.plot(dates, sums, 'r-o', linewidth=2, markersize=8, color='#e74c3c')
        mean_sum = np.mean(sums)
        ax.axhline(y=mean_sum, color='#3498db', linestyle='--', linewidth=2, label=f'均值: {mean_sum:.1f}')
        ax.axhspan(9, 16, alpha=0.2, color='#2ecc71', label='常见区间9-16')
        
        for i, (x, y) in enumerate(zip(dates, sums)):
            ax.annotate(str(y), (x, y), textcoords="offset points", xytext=(0, 10), ha='center', fontsize=10)
        
        ax.set_title('和值走势图（最近15期）', fontsize=16)
        ax.set_xlabel('期数', fontsize=12)
        ax.set_ylabel('和值', fontsize=12)
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
        fig, ax = plt.subplots(figsize=(14, 7))
        dates = [str(i+1) for i in range(len(recent_15))]
        spans = recent_15['跨度'].tolist()
        
        ax.plot(dates, spans, 'g-s', linewidth=2, markersize=8, color='#27ae60')
        mean_span = np.mean(spans)
        ax.axhline(y=mean_span, color='#e67e22', linestyle='--', linewidth=2, label=f'均值: {mean_span:.1f}')
        
        for i, (x, y) in enumerate(zip(dates, spans)):
            ax.annotate(str(y), (x, y), textcoords="offset points", xytext=(0, 10), ha='center', fontsize=10)
        
        ax.set_title('跨度走势图（最近15期）', fontsize=16)
        ax.set_xlabel('期数', fontsize=12)
        ax.set_ylabel('跨度', fontsize=12)
        ax.set_ylim(-0.5, 9.5)
        ax.set_yticks(range(10))
        ax.legend()
        ax.grid(True, alpha=0.3)
        
        plt.tight_layout()
        plt.savefig('fc3d_span_trend.png', dpi=150, bbox_inches='tight')
        plt.close()
        charts.append('fc3d_span_trend.png')
        print("  ✅ 跨度走势图")
    except Exception as e:
        print(f"  ⚠️ 跨度走势图失败: {e}")
    
    # 图3：各位置频率图
    try:
        fig, axes = plt.subplots(1, 3, figsize=(18, 6))
        positions = ['百位', '十位', '个位']
        colors_bar = ['#e74c3c', '#3498db', '#27ae60']
        
        for i, pos in enumerate(positions):
            counts = df[pos].value_counts().sort_index()
            axes[i].bar(counts.index, counts.values, color=colors_bar[i], alpha=0.8, edgecolor='white')
            axes[i].set_title(f'{pos}位', fontsize=14)
            axes[i].set_xlabel('数字', fontsize=11)
            axes[i].set_ylabel('出现次数', fontsize=11)
            axes[i].set_xticks(range(10))
            for bar, v in zip(axes[i].patches, counts.values):
                axes[i].text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5, str(v), ha='center', fontsize=10)
        
        fig.suptitle('各位置数字频率分布', fontsize=16)
        plt.tight_layout()
        plt.savefig('fc3d_position_freq.png', dpi=150, bbox_inches='tight')
        plt.close()
        charts.append('fc3d_position_freq.png')
        print("  ✅ 位置频率图")
    except Exception as e:
        print(f"  ⚠️ 位置频率图失败: {e}")
    
    print(f"\n✅ 共生成 {len(charts)} 张走势图")
    return charts


# ============================================================================
# 分析函数
# ============================================================================
def analyze_pattern(df):
    """分析形态规律"""
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
    """分析各位置"""
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
        
        pos_data[pos] = {
            'top3': top3,
            'miss': miss,
            'counts': total_counts,
            'position_scores': position_scores
        }


def analyze_features(df):
    """综合分析"""
    sums = df['和值'].tolist()
    sum_cnt = Counter(sums)
    
    spans = df['跨度'].tolist()
    span_cnt = Counter(spans)
    
    # 奇偶组合
    parity = []
    for _, r in df.iterrows():
        p = ('奇' if r['百位']%2 else '偶') + ('奇' if r['十位']%2 else '偶') + ('奇' if r['个位']%2 else '偶')
        parity.append(p)
    p_cnt = Counter(parity)
    
    # 大小组合
    size = []
    for _, r in df.iterrows():
        s = ('大' if r['百位']>=5 else '小') + ('大' if r['十位']>=5 else '小') + ('大' if r['个位']>=5 else '小')
        size.append(s)
    s_cnt = Counter(size)
    
    # 和值区间分布
    intervals = [(0, 4), (5, 8), (9, 12), (13, 16), (17, 20), (21, 27)]
    interval_counts = {}
    for low, high in intervals:
        cnt = len([s for s in sums if low <= s <= high])
        interval_counts[f'{low}-{high}'] = cnt
    
    return {
        'common_sum': sum_cnt.most_common(1)[0][0],
        'common_sum_top5': [s for s, _ in sum_cnt.most_common(5)],
        'common_parity': p_cnt.most_common(1)[0][0],
        'common_size': s_cnt.most_common(1)[0][0],
        'span_cnt': span_cnt,
        'p_cnt': p_cnt,
        's_cnt': s_cnt,
        'interval_counts': interval_counts,
        'avg_sum': np.mean(sums)
    }


def predict(df, pos_data, features, predicted_pattern):
    """推演下一期"""
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


# ============================================================================
# 生成HTML报告（完整满血版 + 全端自适应）
# ============================================================================
def image_to_base64(image_path):
    """将图片转换为base64编码"""
    import base64
    if os.path.exists(image_path):
        with open(image_path, 'rb') as f:
            return base64.b64encode(f.read()).decode('utf-8')
    return None


def generate_html_report(df, pos_data, features, predicted_pattern, candidates, next_period, accuracy_result, chart_files):
    """生成HTML报告 - 完整满血版 + 全端自适应"""
    
    recent_5 = df.tail(5)
    
    pos_scores = {}
    for pos in ['百位', '十位', '个位']:
        scores = pos_data[pos]['position_scores']
        sorted_scores = sorted(scores.items(), key=lambda x: x[1], reverse=True)[:5]
        pos_scores[pos] = sorted_scores
    
    group6_pct = features['p_cnt'].get('组六', 0) / len(df) * 100
    group3_pct = features['p_cnt'].get('组三', 0) / len(df) * 100
    group_bz_pct = features['p_cnt'].get('豹子', 0) / len(df) * 100
    
    # 将图片转换为base64嵌入
    sum_img_base64 = image_to_base64('fc3d_sum_trend.png')
    span_img_base64 = image_to_base64('fc3d_span_trend.png')
    pos_img_base64 = image_to_base64('fc3d_position_freq.png')
    
    # 构建图表HTML
    charts_html = ''
    if sum_img_base64:
        charts_html += f'''
                    <div class="chart-card">
                        <h3>📈 和值走势图</h3>
                        <img src="data:image/png;base64,{sum_img_base64}" alt="和值走势图">
                    </div>
'''
    if span_img_base64:
        charts_html += f'''
                    <div class="chart-card">
                        <h3>📐 跨度走势图</h3>
                        <img src="data:image/png;base64,{span_img_base64}" alt="跨度走势图">
                    </div>
'''
    if pos_img_base64:
        charts_html += f'''
                    <div class="chart-card">
                        <h3>📊 位置频率图</h3>
                        <img src="data:image/png;base64,{pos_img_base64}" alt="位置频率图">
                    </div>
'''
    
    acc_html = ""
    if accuracy_result:
        acc = accuracy_result['accuracy']
        acc_html = f'''
        <div class="stat-card">
            <div class="stat-value">{acc:.1f}%</div>
            <div class="stat-label">历史命中率</div>
            <div class="stat-desc">基于27注热号组合 | 测试{accuracy_result['total_tested']}期 命中{accuracy_result['hit_count']}期</div>
        </div>
        '''
    else:
        acc_html = '''
        <div class="stat-card">
            <div class="stat-value">--</div>
            <div class="stat-label">历史命中率</div>
            <div class="stat-desc">数据积累中，继续运行即可显示</div>
        </div>
        '''
    
    html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover, maximum-scale=1.0, user-scalable=yes">
    <title>福彩3D · 智选分析报告</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Microsoft YaHei', 'PingFang SC', 'Helvetica Neue', Arial, sans-serif;
            background: linear-gradient(145deg, #0f172a 0%, #1e293b 100%);
            min-height: 100vh;
            padding: 16px;
        }}
        
        /* 响应式容器 - 全端自适应 */
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            width: 100%;
        }}
        
        /* 头部 */
        .hero {{
            text-align: center;
            margin-bottom: 24px;
        }}
        
        .hero h1 {{
            font-size: clamp(24px, 6vw, 42px);
            font-weight: 800;
            background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
            -webkit-background-clip: text;
            background-clip: text;
            color: transparent;
            margin-bottom: 8px;
        }}
        
        .hero-badge {{
            display: flex;
            flex-wrap: wrap;
            justify-content: center;
            gap: 8px 16px;
            background: rgba(255,255,255,0.1);
            backdrop-filter: blur(10px);
            padding: 8px 16px;
            border-radius: 100px;
            font-size: 12px;
            color: #a5b4fc;
        }}
        
        /* 卡片 - 全端圆角统一 */
        .glass-card {{
            background: rgba(30, 41, 59, 0.7);
            backdrop-filter: blur(12px);
            border-radius: 24px;
            border: 1px solid rgba(255,255,255,0.1);
            margin-bottom: 20px;
            overflow: hidden;
        }}
        
        .card-header {{
            padding: 16px 20px;
            border-bottom: 1px solid rgba(255,255,255,0.08);
            display: flex;
            align-items: center;
            gap: 10px;
            flex-wrap: wrap;
        }}
        
        .card-header h2 {{
            font-size: 18px;
            font-weight: 600;
            color: #f1f5f9;
        }}
        
        .card-body {{
            padding: 20px;
        }}
        
        /* 推演网格 - 手机竖排，平板横排，电脑横排 */
        .prediction-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 16px;
        }}
        
        @media (max-width: 640px) {{
            .prediction-grid {{
                grid-template-columns: 1fr;
            }}
        }}
        
        .prediction-item {{
            background: linear-gradient(145deg, #1e293b, #0f172a);
            border-radius: 28px;
            padding: 20px 16px;
            text-align: center;
            border: 1px solid rgba(255,255,255,0.05);
        }}
        
        .prediction-rank {{
            font-size: 13px;
            color: #94a3b8;
            margin-bottom: 12px;
            letter-spacing: 1px;
        }}
        
        .prediction-balls {{
            display: flex;
            justify-content: center;
            gap: 12px;
            margin-bottom: 12px;
            flex-wrap: wrap;
        }}
        
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
        
        .prediction-score {{
            display: inline-block;
            padding: 4px 14px;
            border-radius: 40px;
            font-size: 13px;
            font-weight: 600;
        }}
        
        .score-high {{ background: linear-gradient(135deg, #f5576c, #f093fb); color: white; }}
        .score-mid {{ background: linear-gradient(135deg, #f59e0b, #f97316); color: white; }}
        .score-low {{ background: linear-gradient(135deg, #10b981, #34d399); color: white; }}
        
        .prediction-detail {{
            font-size: 11px;
            color: #94a3b8;
            margin-top: 8px;
        }}
        
        /* 统计卡片网格 - 响应式 */
        .stats-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
            gap: 16px;
            margin-bottom: 20px;
        }}
        
        @media (max-width: 640px) {{
            .stats-grid {{
                grid-template-columns: repeat(2, 1fr);
            }}
        }}
        
        @media (max-width: 400px) {{
            .stats-grid {{
                grid-template-columns: 1fr;
            }}
        }}
        
        .stat-card {{
            background: rgba(30, 41, 59, 0.6);
            border-radius: 20px;
            padding: 16px;
            text-align: center;
            border: 1px solid rgba(255,255,255,0.05);
        }}
        
        .stat-value {{
            font-size: clamp(28px, 8vw, 36px);
            font-weight: 800;
            color: #f1f5f9;
            margin-bottom: 6px;
        }}
        
        .stat-label {{
            font-size: 12px;
            color: #94a3b8;
        }}
        
        .stat-desc {{
            font-size: 10px;
            color: #64748b;
            margin-top: 4px;
        }}
        
        /* 表格 - 横向滚动适配手机 */
        .table-wrapper {{
            overflow-x: auto;
            -webkit-overflow-scrolling: touch;
        }}
        
        .data-table {{
            width: 100%;
            border-collapse: collapse;
            min-width: 500px;
        }}
        
        .data-table th {{
            text-align: left;
            padding: 12px 10px;
            color: #94a3b8;
            font-weight: 500;
            font-size: 12px;
            border-bottom: 1px solid rgba(255,255,255,0.08);
        }}
        
        .data-table td {{
            padding: 12px 10px;
            color: #e2e8f0;
            border-bottom: 1px solid rgba(255,255,255,0.05);
            font-size: 13px;
        }}
        
        /* 号码球 - 响应式大小 */
        .ball-small {{
            display: inline-flex;
            width: clamp(32px, 10vw, 38px);
            height: clamp(32px, 10vw, 38px);
            background: linear-gradient(145deg, #f093fb, #f5576c);
            border-radius: 50%;
            align-items: center;
            justify-content: center;
            font-weight: 700;
            font-size: clamp(14px, 4vw, 16px);
            color: white;
            margin: 0 2px;
        }}
        
        .ball-blue {{ background: linear-gradient(145deg, #4facfe, #00f2fe); }}
        .ball-green {{ background: linear-gradient(145deg, #43e97b, #38f9d7); }}
        .ball-purple {{ background: linear-gradient(145deg, #a18cd1, #fbc2eb); }}
        
        /* 进度条 */
        .progress-bar {{
            background: rgba(255,255,255,0.1);
            border-radius: 12px;
            height: 32px;
            overflow: hidden;
            margin: 12px 0;
        }}
        
        .progress-fill {{
            background: linear-gradient(90deg, #f093fb, #f5576c);
            height: 100%;
            display: flex;
            align-items: center;
            padding-left: 12px;
            color: white;
            font-size: 12px;
            font-weight: 500;
            border-radius: 12px;
        }}
        
        /* 图表网格 - 手机竖排 */
        .charts-grid {{
            display: flex;
            flex-direction: column;
            gap: 24px;
        }}
        
        .chart-card {{
            background: rgba(15, 23, 42, 0.8);
            border-radius: 20px;
            padding: 16px;
            text-align: center;
        }}
        
        .chart-card h3 {{
            color: #f1f5f9;
            margin-bottom: 12px;
            font-size: 16px;
        }}
        
        .chart-card img {{
            width: 100%;
            height: auto;
            border-radius: 12px;
        }}
        
        /* 备选参考网格 */
        .candidates-grid {{
            display: flex;
            flex-wrap: wrap;
            gap: 12px;
            justify-content: center;
        }}
        
        .candidate-item {{
            background: rgba(255,255,255,0.05);
            border-radius: 20px;
            padding: 10px 16px;
            text-align: center;
        }}
        
        /* 评分说明网格 */
        .score-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
            gap: 10px;
        }}
        
        .badge-primary {{
            background: rgba(245, 87, 108, 0.15);
            color: #f093fb;
            padding: 4px 10px;
            border-radius: 100px;
            font-size: 11px;
            display: inline-block;
        }}
        
        .footer {{
            text-align: center;
            padding: 24px;
            color: #475569;
            font-size: 11px;
        }}
        
        /* 颜色类 */
        .text-white {{ color: white !important; font-weight: bold; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="hero">
            <h1>🎯 福彩3D · 智选分析</h1>
            <div class="hero-badge">
                <span>📅 {datetime.now().strftime('%Y-%m-%d')}</span>
                <span>⚡ 下一期: {next_period}</span>
                <span>📊 基于 {len(df)} 期历史数据</span>
            </div>
        </div>
        
        <!-- 推演结果 -->
        <div class="glass-card">
            <div class="card-header">
                <span style="font-size: 24px;">⭐</span>
                <h2>智能推演 · 下一期预测</h2>
            </div>
            <div class="card-body">
                <div class="prediction-grid">
'''
    
    for i, c in enumerate(candidates[:5], 1):
        nums = list(c['号码'])
        score_class = "score-high" if c['得分'] >= 16 else "score-mid" if c['得分'] >= 12 else "score-low"
        html += f'''
                    <div class="prediction-item">
                        <div class="prediction-rank">第{i}注 · 推荐</div>
                        <div class="prediction-balls">
                            <div class="ball-large">{nums[0]}</div>
                            <div class="ball-large">{nums[1]}</div>
                            <div class="ball-large">{nums[2]}</div>
                        </div>
                        <div><span class="prediction-score {score_class}">{c['得分']}分</span></div>
                        <div class="prediction-detail">形态: {c['形态']} | 和值: {c['和值']}<br>{" ".join(c['原因'][:2])}</div>
                    </div>
'''
    
    html += f'''
                </div>
            </div>
        </div>
        
        <!-- 统计卡片 -->
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-value">{features['common_sum']}</div>
                <div class="stat-label">🎯 最常见和值</div>
                <div class="stat-desc">重点关注</div>
            </div>
            <div class="stat-card">
                <div class="stat-value">{features['span_cnt'].most_common(1)[0][0]}</div>
                <div class="stat-label">📐 最常见跨度</div>
                <div class="stat-desc">历史最高频</div>
            </div>
            <div class="stat-card">
                <div class="stat-value">{predicted_pattern}</div>
                <div class="stat-label">🔮 预测形态</div>
                <div class="stat-desc">基于转移概率</div>
            </div>
            {acc_html}
        </div>
        
        <!-- 最近5期开奖记录 -->
        <div class="glass-card">
            <div class="card-header">
                <span style="font-size: 24px;">📋</span>
                <h2>最近5期开奖记录</h2>
            </div>
            <div class="card-body">
                <div class="table-wrapper">
                    <table class="data-table">
                        <thead>
                            <tr><th>期号</th><th>开奖日期</th><th>开奖号码</th><th>形态</th><th>和值</th><th>跨度</th></tr>
                        </thead>
                        <tbody>
'''
    
    for _, row in recent_5.iterrows():
        nums = row['开奖号码']
        html += f'''
                            <tr>
                                <td>{row['期号']}</td>
                                <td>{row['开奖日期'].strftime('%Y-%m-%d')}</td>
                                <td><span class="ball-small">{nums[0]}</span><span class="ball-small">{nums[1]}</span><span class="ball-small">{nums[2]}</span></td>
                                <td>{row['形态']}</td>
                                <td>{row['和值']}</td>
                                <td>{row['跨度']}</td>
                            </tr>
'''
    
    html += f'''
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        
        <!-- 形态分布 + 奇偶大小 -->
        <div class="stats-grid" style="grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));">
            <div class="glass-card" style="margin:0;">
                <div class="card-header"><span>🎨</span><h2>形态分布</h2></div>
                <div class="card-body">
                    <div class="progress-bar"><div class="progress-fill" style="width: {group6_pct}%;">组六 {features['p_cnt'].get('组六',0)}期</div></div>
                    <div class="progress-bar"><div class="progress-fill" style="background: linear-gradient(90deg, #f59e0b, #f97316); width: {group3_pct}%;">组三 {features['p_cnt'].get('组三',0)}期</div></div>
                    <div class="progress-bar"><div class="progress-fill" style="background: linear-gradient(90deg, #10b981, #34d399); width: {group_bz_pct}%;">豹子 {features['p_cnt'].get('豹子',0)}期</div></div>
                </div>
            </div>
            <div class="glass-card" style="margin:0;">
                <div class="card-header"><span>🔢</span><h2>奇偶 & 大小</h2></div>
                <div class="card-body">
                    <div><span class="badge-primary">最常见奇偶</span> <strong style="color:#f1f5f9;">{features['common_parity']}</strong></div>
                    <div style="margin-top: 12px;"><span class="badge-primary">最常见大小</span> <strong style="color:#f1f5f9;">{features['common_size']}</strong></div>
                    <div style="margin-top: 16px;"><span class="badge-primary">和值Top5</span> <span style="color:#a5b4fc;">{features['common_sum_top5']}</span></div>
                    <div style="margin-top: 12px;"><span class="badge-primary">平均和值</span> <span style="color:#a5b4fc;">{features['avg_sum']:.1f}</span></div>
                </div>
            </div>
        </div>
        
        <!-- 三位置分析 -->
        <div class="stats-grid" style="grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));">
'''
    
    for pos, color in [('百位', '#ff6b6b'), ('十位', '#45b7d1'), ('个位', '#96ceb4')]:
        html += f'''
            <div class="glass-card" style="margin:0;">
                <div class="card-header"><span>📍</span><h2>{pos}</h2></div>
                <div class="card-body">
                    <div><span class="badge-primary">热号Top3</span> <strong style="color:#f1f5f9; font-size: 18px;">{pos_data[pos]['top3']}</strong></div>
                    <div style="margin-top: 16px; font-size: 12px; color: #94a3b8;">合理性评分Top5:</div>
                    <div style="margin-top: 8px;">
'''
        for num, score in pos_scores[pos][:5]:
            html += f'<span class="ball-small" style="background: {color}; width: 32px; height: 32px; font-size: 14px;">{num}</span> '
        html += f'''
                    </div>
                </div>
            </div>
'''
    
    html += f'''
        </div>
        
        <!-- 走势图 -->
        <div class="glass-card">
            <div class="card-header">
                <span>📈</span>
                <h2>走势图分析（点击图片可放大）</h2>
            </div>
            <div class="card-body">
                <div class="charts-grid">
                    {charts_html}
                </div>
            </div>
        </div>
        
        <!-- 备选参考 -->
        <div class="glass-card">
            <div class="card-header">
                <span>🔄</span>
                <h2>备选参考 · 形态匹配组合</h2>
            </div>
            <div class="card-body">
                <div class="candidates-grid">
'''
    
    pattern_matched = [c for c in candidates if c['形态'] == predicted_pattern and c not in candidates[:5]]
    for c in pattern_matched[:12]:
        nums = list(c['号码'])
        html += f'''
                    <div class="candidate-item">
                        <div><span class="ball-small">{nums[0]}</span><span class="ball-small">{nums[1]}</span><span class="ball-small">{nums[2]}</span></div>
                        <div style="font-size: 11px; color: #94a3b8; margin-top: 4px;">{c['得分']}分</div>
                    </div>
'''
    
    html += f'''
                </div>
            </div>
        </div>
        
        <!-- 评分说明 -->
        <div class="glass-card">
            <div class="card-header">
                <span>📖</span>
                <h2>评分系统 · 满分20分</h2>
            </div>
            <div class="card-body">
                <div class="score-grid">
                    <div><span class="badge-primary">🎯 和值精确</span> <span class="text-white">+5分</span></div>
                    <div><span class="badge-primary">📊 和值接近</span> <span class="text-white">+2分</span></div>
                    <div><span class="badge-primary">✨ 奇偶匹配</span> <span class="text-white">+4分</span></div>
                    <div><span class="badge-primary">📏 大小匹配</span> <span class="text-white">+4分</span></div>
                    <div><span class="badge-primary">🔮 形态匹配</span> <span class="text-white">+3分</span></div>
                    <div><span class="badge-primary">📍 位置合理性</span> <span class="text-white">+3分</span></div>
                    <div><span class="badge-primary">🛡️ 防重奖励</span> <span class="text-white">+1分</span></div>
                </div>
            </div>
        </div>
        
        <div class="footer">
            ⚠️ 本分析基于历史数据统计规律，仅供学习参考<br>
            彩票开奖是独立随机事件，任何号码中奖概率相同
        </div>
    </div>
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
    
    candidates = predict(df_recent, pos_data, features, predicted_pattern)
    
    accuracy_result = calculate_accuracy(df)
    
    chart_files = generate_charts(df)
    
    html_file = generate_html_report(df_recent, pos_data, features, predicted_pattern, candidates, next_period, accuracy_result, chart_files)
    
    print(f"\n✅ HTML报告已生成: {html_file}")
    print(f"📁 位置: {os.path.abspath(html_file)}")


if __name__ == "__main__":
    main()
