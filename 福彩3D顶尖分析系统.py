def fetch_data():
    """抓取福彩3D历史数据 - 使用500.com API"""
    all_data = []
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Accept': 'application/json',
        'Referer': 'https://www.500.com/'
    }
    
    # 使用500.com的公开API
    api_url = 'https://www.500.com/api/proxy/lottery?lotteryId=1&action=history&pageSize=200'
    
    try:
        r = requests.get(api_url, headers=headers, timeout=15)
        if r.status_code == 200:
            data = r.json()
            if 'data' in data and 'list' in data['data']:
                for item in data['data']['list']:
                    period = str(item.get('period', ''))
                    date = item.get('date', '')
                    red = item.get('red', '')
                    if red and ',' in red:
                        reds = red.split(',')
                        if len(reds) >= 3:
                            nums = f"{reds[0]}{reds[1]}{reds[2]}"
                            all_data.append({'期号': period, '开奖日期': date, '开奖号码': nums})
    except Exception as e:
        print(f"  API抓取失败: {e}")
    
    if not all_data:
        # 备用：直接解析HTML
        try:
            url = 'https://www.500.com/cp/fc3d/kaijiang/'
            r = requests.get(url, headers=headers, timeout=15)
            soup = BeautifulSoup(r.text, 'html.parser')
            
            # 查找表格数据
            rows = soup.find_all('tr')
            for row in rows[1:21]:
                cols = row.find_all('td')
                if len(cols) >= 3:
                    period = cols[0].text.strip()
                    date = cols[1].text.strip()
                    em = row.find_all('em')
                    if len(em) >= 3:
                        nums = f"{em[0].text.strip()}{em[1].text.strip()}{em[2].text.strip()}"
                        all_data.append({'期号': period, '开奖日期': date, '开奖号码': nums})
        except Exception as e:
            print(f"  备用抓取失败: {e}")
    
    # 去重排序
    seen = set()
    unique = []
    for d in all_data:
        if d['期号'] not in seen:
            seen.add(d['期号'])
            unique.append(d)
    unique.sort(key=lambda x: int(x['期号']), reverse=True)
    
    if unique:
        print(f"  ✅ 成功抓取 {len(unique)} 期数据，最新期号: {unique[0]['期号']}")
    else:
        print("  ⚠️ 抓取失败，未获取到数据")
    
    return unique
