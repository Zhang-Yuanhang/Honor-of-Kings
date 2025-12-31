import pandas as pd
import numpy as np
from collections import defaultdict
import warnings
import os
from datetime import datetime
import base64

warnings.filterwarnings('ignore')

# 创建示例数据
data = pd.read_excel('C:/Files/Ubiquant/code/HOK/hok_bp/practicing/内战data/内战计分表.xlsx')

# 创建列名
columns = ['比赛ID', '比赛时间', '胜方', '蓝方边', '蓝方野', '蓝方中', '蓝方射', '蓝方辅', '蓝方MVP', 
           '红方边', '红方野', '红方中', '红方射', '红方辅', '红方MVP']

# 创建DataFrame
df = data[columns].copy()

print("正在处理比赛数据...")
print(f"数据规模：{len(df)}场比赛，{len(columns)}列")
print("\n" + "="*80 + "\n")

# 定义分路映射
position_map = {
    '蓝方边': '边路', '蓝方野': '打野', '蓝方中': '中路', '蓝方射': '发育路', '蓝方辅': '游走',
    '红方边': '边路', '红方野': '打野', '红方中': '中路', '红方射': '发育路', '红方辅': '游走'
}

# MVP位置映射
mvp_position_map = {
    '边': '边路', '野': '打野', '中': '中路', '射': '发育路', '辅': '游走'
}

# 初始化统计字典
player_stats = defaultdict(lambda: {
    '总场次': 0, '总胜场': 0,
    '边路场次': 0, '边路胜场': 0,
    '打野场次': 0, '打野胜场': 0,
    '中路场次': 0, '中路胜场': 0,
    '发育路场次': 0, '发育路胜场': 0,
    '游走场次': 0, '游走胜场': 0,
    'MVP次数': 0,
    '英雄池': set(),
    '边路英雄池': set(),
    '打野英雄池': set(),
    '中路英雄池': set(),
    '发育路英雄池': set(),
    '游走英雄池': set(),
    '英雄胜场': defaultdict(int),  # 新增：统计每个英雄的胜场
    '英雄场次': defaultdict(int),   # 新增：统计每个英雄的场次
})

hero_stats = defaultdict(lambda: {
    '总场次': 0, '总胜场': 0,
    '边路场次': 0, '边路胜场': 0,
    '打野场次': 0, '打野胜场': 0,
    '中路场次': 0, '中路胜场': 0,
    '发育路场次': 0, '发育路胜场': 0,
    '游走场次': 0, '游走胜场': 0,
    '玩家胜场': defaultdict(int),  # 新增：统计每个玩家的胜场
    '玩家场次': defaultdict(int),   # 新增：统计每个玩家的场次
})

# 处理每场比赛
for idx, row in df.iterrows():
    winner = row['胜方']
    match_id = row['比赛ID']
    
    # 处理蓝方
    for pos_col, position in position_map.items():
        if pos_col.startswith('蓝方'):
            cell_value = row[pos_col]
            if isinstance(cell_value, str) and '-' in cell_value:
                player, hero = cell_value.split('-', 1)
                
                # 更新玩家统计
                player_stats[player]['总场次'] += 1
                player_stats[player][f'{position}场次'] += 1
                
                # 更新英雄池
                player_stats[player]['英雄池'].add(hero)
                player_stats[player][f'{position}英雄池'].add(hero)
                
                # 更新玩家-英雄统计
                player_stats[player]['英雄场次'][hero] += 1
                
                if winner == '蓝':
                    player_stats[player]['总胜场'] += 1
                    player_stats[player][f'{position}胜场'] += 1
                    player_stats[player]['英雄胜场'][hero] += 1
                
                # 更新英雄统计
                hero_stats[hero]['总场次'] += 1
                hero_stats[hero][f'{position}场次'] += 1
                
                # 更新英雄-玩家统计
                hero_stats[hero]['玩家场次'][player] += 1
                
                if winner == '蓝':
                    hero_stats[hero]['总胜场'] += 1
                    hero_stats[hero][f'{position}胜场'] += 1
                    hero_stats[hero]['玩家胜场'][player] += 1
                
                # 检查MVP
                mvp_pos = row['蓝方MVP']
                if mvp_pos in mvp_position_map and mvp_position_map[mvp_pos] == position:
                    player_stats[player]['MVP次数'] += 1
    
    # 处理红方
    for pos_col, position in position_map.items():
        if pos_col.startswith('红方'):
            cell_value = row[pos_col]
            if isinstance(cell_value, str) and '-' in cell_value:
                player, hero = cell_value.split('-', 1)
                
                # 更新玩家统计
                player_stats[player]['总场次'] += 1
                player_stats[player][f'{position}场次'] += 1
                
                # 更新英雄池
                player_stats[player]['英雄池'].add(hero)
                player_stats[player][f'{position}英雄池'].add(hero)
                
                # 更新玩家-英雄统计
                player_stats[player]['英雄场次'][hero] += 1
                
                if winner == '红':
                    player_stats[player]['总胜场'] += 1
                    player_stats[player][f'{position}胜场'] += 1
                    player_stats[player]['英雄胜场'][hero] += 1
                
                # 更新英雄统计
                hero_stats[hero]['总场次'] += 1
                hero_stats[hero][f'{position}场次'] += 1
                
                # 更新英雄-玩家统计
                hero_stats[hero]['玩家场次'][player] += 1
                
                if winner == '红':
                    hero_stats[hero]['总胜场'] += 1
                    hero_stats[hero][f'{position}胜场'] += 1
                    hero_stats[hero]['玩家胜场'][player] += 1
                
                # 检查MVP
                mvp_pos = row['红方MVP']
                if mvp_pos in mvp_position_map and mvp_position_map[mvp_pos] == position:
                    player_stats[player]['MVP次数'] += 1

# 1. 玩家总场次+胜率排行榜
def create_player_leaderboard():
    leaderboard = []
    for player, stats in player_stats.items():
        if stats['总场次'] > 0:
            win_rate = stats['总胜场'] / stats['总场次']
            leaderboard.append({
                '玩家': player,
                '总场次': stats['总场次'],
                '总胜场': stats['总胜场'],
                '总胜率': win_rate,
                '总胜率百分比': f"{win_rate * 100:.2f}%",
                'MVP次数': stats['MVP次数'],
                '英雄池数量': len(stats['英雄池'])
            })
    
    leaderboard_df = pd.DataFrame(leaderboard)
    leaderboard_df = leaderboard_df.sort_values(by=['总胜率', '总场次'], ascending=[False, False])
    leaderboard_df = leaderboard_df.reset_index(drop=True)
    leaderboard_df.index = leaderboard_df.index + 1
    leaderboard_df.index.name = '排名'
    
    return leaderboard_df

# 2. 英雄总场次+胜率排行榜
def create_hero_leaderboard():
    leaderboard = []
    for hero, stats in hero_stats.items():
        if stats['总场次'] > 0:
            win_rate = stats['总胜场'] / stats['总场次']
            leaderboard.append({
                '英雄': hero,
                '总场次': stats['总场次'],
                '总胜场': stats['总胜场'],
                '总胜率': win_rate,
                '总胜率百分比': f"{win_rate * 100:.2f}%"
            })
    
    leaderboard_df = pd.DataFrame(leaderboard)
    leaderboard_df = leaderboard_df.sort_values(by=['总胜率', '总场次'], ascending=[False, False])
    leaderboard_df = leaderboard_df.reset_index(drop=True)
    leaderboard_df.index = leaderboard_df.index + 1
    leaderboard_df.index.name = '排名'
    
    return leaderboard_df

# 3. 玩家MVP榜
def create_mvp_leaderboard():
    leaderboard = []
    for player, stats in player_stats.items():
        if stats['MVP次数'] > 0:
            leaderboard.append({
                '玩家': player,
                'MVP次数': stats['MVP次数']
            })
    
    leaderboard_df = pd.DataFrame(leaderboard)
    leaderboard_df = leaderboard_df.sort_values(by='MVP次数', ascending=False)
    leaderboard_df = leaderboard_df.reset_index(drop=True)
    leaderboard_df.index = leaderboard_df.index + 1
    leaderboard_df.index.name = '排名'
    
    return leaderboard_df

# 4. 玩家英雄池数量排行榜
def create_hero_pool_leaderboard():
    leaderboard = []
    for player, stats in player_stats.items():
        if stats['总场次'] > 0:
            hero_pool_size = len(stats['英雄池'])
            
            leaderboard.append({
                '玩家': player,
                '英雄池数量': hero_pool_size,
                '总场次': stats['总场次'],
                '平均每英雄场次': round(stats['总场次'] / hero_pool_size, 2) if hero_pool_size > 0 else 0,
            })
    
    leaderboard_df = pd.DataFrame(leaderboard)
    leaderboard_df = leaderboard_df.sort_values(by=['英雄池数量', '总场次'], ascending=[False, False])
    leaderboard_df = leaderboard_df.reset_index(drop=True)
    leaderboard_df.index = leaderboard_df.index + 1
    leaderboard_df.index.name = '排名'
    
    return leaderboard_df

# 5. 各分路玩家排行榜
def create_position_player_leaderboard(position):
    position_name = position
    leaderboard = []
    
    for player, stats in player_stats.items():
        games = stats[f'{position_name}场次']
        if games > 0:
            wins = stats[f'{position_name}胜场']
            win_rate = wins / games
            hero_pool_size = len(stats.get(f'{position_name}英雄池', set()))
            
            leaderboard.append({
                '玩家': player,
                '场次': games,
                '胜场': wins,
                '胜率': win_rate,
                '胜率百分比': f"{win_rate * 100:.0f}%",
                f'{position_name}英雄池': hero_pool_size,
                f'平均每英雄场次': round(games / hero_pool_size, 2) if hero_pool_size > 0 else 0
            })
    
    leaderboard_df = pd.DataFrame(leaderboard)
    leaderboard_df = leaderboard_df.sort_values(by=['胜率', '场次'], ascending=[False, False])
    leaderboard_df = leaderboard_df.reset_index(drop=True)
    leaderboard_df.index = leaderboard_df.index + 1
    leaderboard_df.index.name = '排名'
    
    return leaderboard_df

# 6. 各分路英雄排行榜
def create_position_hero_leaderboard(position):
    position_name = position
    leaderboard = []
    
    for hero, stats in hero_stats.items():
        games = stats[f'{position_name}场次']
        if games > 0:
            wins = stats[f'{position_name}胜场']
            win_rate = wins / games
            leaderboard.append({
                '英雄': hero,
                '场次': games,
                '胜场': wins,
                '胜率': win_rate,
                '胜率百分比': f"{win_rate * 100:.2f}%"
            })
    
    leaderboard_df = pd.DataFrame(leaderboard)
    leaderboard_df = leaderboard_df.sort_values(by=['胜率', '场次'], ascending=[False, False])
    leaderboard_df = leaderboard_df.reset_index(drop=True)
    leaderboard_df.index = leaderboard_df.index + 1
    leaderboard_df.index.name = '排名'
    
    return leaderboard_df

# 7. 玩家分路多样性分析
def create_player_position_diversity():
    leaderboard = []
    for player, stats in player_stats.items():
        if stats['总场次'] > 0:
            # 统计玩家打过的分路数量
            played_positions = 0
            position_list = []
            for position in ['边路', '打野', '中路', '发育路', '游走']:
                if stats[f'{position}场次'] > 0:
                    played_positions += 1
                    position_list.append(position)
            
            # 计算分路专注度
            max_position_games = 0
            main_position = '无'
            for position in ['边路', '打野', '中路', '发育路', '游走']:
                if stats[f'{position}场次'] > max_position_games:
                    max_position_games = stats[f'{position}场次']
                    main_position = position
            
            position_concentration = max_position_games / stats['总场次'] if stats['总场次'] > 0 else 0
            
            leaderboard.append({
                '玩家': player,
                '总场次': stats['总场次'],
                '使用分路数': played_positions,
                '使用分路': ', '.join(position_list),
                '主要分路': main_position,
                '主要分路场次': max_position_games,
                '分路专注度': position_concentration,
                '分路专注度百分比': f"{position_concentration * 100:.1f}%",
                '英雄池数量': len(stats['英雄池'])
            })
    
    leaderboard_df = pd.DataFrame(leaderboard)
    leaderboard_df = leaderboard_df.sort_values(by=['使用分路数', '英雄池数量'], ascending=[False, False])
    leaderboard_df = leaderboard_df.reset_index(drop=True)
    leaderboard_df.index = leaderboard_df.index + 1
    leaderboard_df.index.name = '排名'
    
    return leaderboard_df

# 8. 同一个英雄，玩家胜率榜
def create_hero_player_winrate_leaderboard():
    """生成每个英雄的玩家胜率排行榜"""
    hero_player_stats = {}
    
    for hero, stats in hero_stats.items():
        player_list = []
        for player in stats['玩家场次']:
            games = stats['玩家场次'][player]
            wins = stats['玩家胜场'][player]
            if games > 0:
                win_rate = wins / games
                player_list.append({
                    '玩家': player,
                    '场次': games,
                    '胜场': wins,
                    '胜率': win_rate,
                    '胜率百分比': f"{win_rate * 100:.2f}%"
                })
        
        # 按胜率排序
        if player_list:
            player_list.sort(key=lambda x: x['胜率'], reverse=True)
            hero_player_stats[hero] = player_list[:5]  # 只取前5名
    
    return hero_player_stats

# 9. 同一个玩家，英雄胜率榜
def create_player_hero_winrate_leaderboard():
    """生成每个玩家的英雄胜率排行榜"""
    player_hero_stats = {}
    
    for player, stats in player_stats.items():
        hero_list = []
        for hero in stats['英雄场次']:
            games = stats['英雄场次'][hero]
            wins = stats['英雄胜场'][hero]
            if games > 0:
                win_rate = wins / games
                hero_list.append({
                    '英雄': hero,
                    '场次': games,
                    '胜场': wins,
                    '胜率': win_rate,
                    '胜率百分比': f"{win_rate * 100:.2f}%"
                })
        
        # 按胜率排序
        if hero_list:
            hero_list.sort(key=lambda x: x['胜率'], reverse=True)
            player_hero_stats[player] = hero_list[:5]  # 只取前5名
    
    return player_hero_stats

# 生成所有排行榜
print("正在生成排行榜数据...")
player_leaderboard = create_player_leaderboard()
hero_leaderboard = create_hero_leaderboard()
mvp_leaderboard = create_mvp_leaderboard()
hero_pool_leaderboard = create_hero_pool_leaderboard()
position_diversity = create_player_position_diversity()

positions = ['边路', '打野', '中路', '发育路', '游走']
position_leaderboards = {}
for position in positions:
    position_leaderboards[position] = {
        'player': create_position_player_leaderboard(position),
        'hero': create_position_hero_leaderboard(position)
    }

# 生成新增的排行榜
hero_player_leaderboard = create_hero_player_winrate_leaderboard()
player_hero_leaderboard = create_player_hero_winrate_leaderboard()

# 生成可视化图表
print("正在生成可视化图表...")
try:
    import matplotlib.pyplot as plt
    import seaborn as sns
    
    # 设置中文字体
    plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'DejaVu Sans']
    plt.rcParams['axes.unicode_minus'] = False
    
    # 创建图表保存目录
    charts_dir = "charts"
    if not os.path.exists(charts_dir):
        os.makedirs(charts_dir)
    
    # 1. 胜率TOP10玩家图表
    fig1, ax1 = plt.subplots(figsize=(10, 6))
    top_players = player_leaderboard.head(10).copy()
    top_players = top_players.sort_values(by='总胜率',ascending=False)
    colors = plt.cm.viridis(np.linspace(0.2, 0.8, len(top_players)))
    bars = ax1.barh(top_players['玩家'], top_players['总胜率'], color=colors, edgecolor='black')
    ax1.set_xlabel('胜率', fontsize=10)
    ax1.set_title('胜率TOP10玩家', fontsize=12, fontweight='bold')
    ax1.set_xlim(0, 1)
    ax1.invert_yaxis()
    
    # 添加数值标签
    for i, (bar, row) in enumerate(zip(bars, top_players.itertuples())):
        ax1.text(bar.get_width() + 0.01, bar.get_y() + bar.get_height()/2,
                f'{row.总胜率:.2%}', ha='left', va='center', fontsize=9)
    
    plt.tight_layout()
    plt.savefig(f'{charts_dir}/胜率TOP10玩家.png', dpi=120, bbox_inches='tight')
    plt.close()
    
    # 2. 英雄池数量TOP10图表
    fig2, ax2 = plt.subplots(figsize=(10, 6))
    hero_pool_top10 = hero_pool_leaderboard.head(10).copy()
    hero_pool_top10 = hero_pool_top10.sort_values(by='英雄池数量',ascending=False)
    colors = plt.cm.plasma(np.linspace(0.2, 0.8, len(hero_pool_top10)))
    bars = ax2.barh(hero_pool_top10['玩家'], hero_pool_top10['英雄池数量'], color=colors, edgecolor='black')
    ax2.set_xlabel('英雄池数量', fontsize=10)
    ax2.set_title('英雄池数量TOP10玩家', fontsize=12, fontweight='bold')
    ax2.invert_yaxis()
    
    # 添加数值标签
    for bar in bars:
        width = bar.get_width()
        ax2.text(width + 0.1, bar.get_y() + bar.get_height()/2,
                f'{int(width)}', ha='left', va='center', fontsize=9)
    
    plt.tight_layout()
    plt.savefig(f'{charts_dir}/英雄池数量TOP10.png', dpi=120, bbox_inches='tight')
    plt.close()
    
    # 3. MVP次数分布图表
    fig3, ax3 = plt.subplots(figsize=(8, 5))
    mvp_data = mvp_leaderboard.copy()
    if len(mvp_data) > 0:
        mvp_data = mvp_data.sort_values(by='MVP次数',ascending=False)
        colors = ['gold' if i == 0 else 'lightblue' for i in range(len(mvp_data))]
        bars = ax3.barh(mvp_data['玩家'], mvp_data['MVP次数'], color=colors, edgecolor='black')
        ax3.set_xlabel('MVP次数', fontsize=10)
        ax3.set_title('MVP次数排行榜', fontsize=12, fontweight='bold')
        ax3.invert_yaxis()
        
        # 添加数值标签
        for bar in bars:
            width = bar.get_width()
            ax3.text(width + 0.05, bar.get_y() + bar.get_height()/2,
                    f'{int(width)}', ha='left', va='center', fontsize=9)
    
    plt.tight_layout()
    plt.savefig(f'{charts_dir}/MVP次数分布.png', dpi=120, bbox_inches='tight')
    plt.close()
    
    # 4. 英雄胜率TOP10图表
    fig4, ax4 = plt.subplots(figsize=(10, 6))
    top_heroes = hero_leaderboard[hero_leaderboard['总场次'] >= 2].head(10).copy()
    if len(top_heroes) > 0:
        top_heroes = top_heroes.sort_values(by='总胜率',ascending=False)
        colors = plt.cm.Set3(np.linspace(0.1, 0.9, len(top_heroes)))
        bars = ax4.barh(top_heroes['英雄'], top_heroes['总胜率'], color=colors, edgecolor='black')
        ax4.set_xlabel('胜率', fontsize=10)
        ax4.set_title('英雄胜率TOP10（出场≥2次）', fontsize=12, fontweight='bold')
        ax4.set_xlim(0, 1)
        ax4.invert_yaxis()
        
        # 添加数值标签
        for i, (bar, row) in enumerate(zip(bars, top_heroes.itertuples())):
            ax4.text(bar.get_width() + 0.01, bar.get_y() + bar.get_height()/2,
                    f'{row.总胜率:.2%}', ha='left', va='center', fontsize=8)
    
    plt.tight_layout()
    plt.savefig(f'{charts_dir}/英雄胜率TOP10.png', dpi=120, bbox_inches='tight')
    plt.close()
    
    # 5. 分路多样性饼图
    fig5, ax5 = plt.subplots(figsize=(8, 6))
    diversity_counts = position_diversity['使用分路数'].value_counts().sort_index()
    colors = plt.cm.Pastel1(np.arange(len(diversity_counts)) / len(diversity_counts))
    wedges, texts, autotexts = ax5.pie(diversity_counts.values, 
                                      labels=[f'{k}个分路' for k in diversity_counts.index],
                                      autopct='%1.1f%%', colors=colors, startangle=90,
                                      textprops={'fontsize': 9})
    ax5.set_title('玩家分路多样性分布', fontsize=12, fontweight='bold')
    plt.tight_layout()
    plt.savefig(f'{charts_dir}/分路多样性分布.png', dpi=120, bbox_inches='tight')
    plt.close()
    
    # 6. 红蓝方胜率对比
    fig6, ax6 = plt.subplots(figsize=(6, 5))
    blue_wins = len(df[df['胜方'] == '蓝'])
    red_wins = len(df[df['胜方'] == '红'])
    total_games = len(df)
    
    colors = ['#1f77b4', '#ff7f0e']
    bars = ax6.bar(['蓝方', '红方'], [blue_wins, red_wins], color=colors, edgecolor='black')
    ax6.set_ylabel('胜场数', fontsize=10)
    ax6.set_title(f'红蓝方胜场对比', fontsize=12, fontweight='bold')
    
    # 添加数值和百分比标签
    for bar, wins in zip(bars, [blue_wins, red_wins]):
        height = bar.get_height()
        percentage = (wins / total_games * 100) if total_games > 0 else 0
        ax6.text(bar.get_x() + bar.get_width()/2, height + 0.1,
                f'{wins}场\n({percentage:.1f}%)', ha='center', va='bottom', fontsize=9)
    
    plt.tight_layout()
    plt.savefig(f'{charts_dir}/红蓝方胜率对比.png', dpi=120, bbox_inches='tight')
    plt.close()
    
    print(f"图表已保存到 {charts_dir} 目录")
    
except ImportError:
    print("警告: matplotlib库未安装，跳过图表生成")
    charts_dir = None
except Exception as e:
    print(f"图表生成错误: {e}")
    charts_dir = None

# 生成HTML报告
def generate_html_report():
    """生成完整的单页HTML报告"""
    
    # 获取当前时间
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # 统计摘要
    total_days = len(df['比赛时间'].unique())
    total_games = len(df)
    total_players = len(player_stats)
    total_heroes = len(hero_stats)
    blue_wins = len(df[df['胜方'] == '蓝'])
    red_wins = len(df[df['胜方'] == '红'])
    
    # 获取关键数据
    top_players = player_leaderboard.head(10)
    top_heroes = hero_leaderboard.head(10)
    top_mvp = mvp_leaderboard.head(10)
    top_hero_pool = hero_pool_leaderboard.head(10)
    
    # 创建HTML内容
    html_content = f"""
    <!DOCTYPE html>
    <html lang="zh-CN">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>🏆 “饼干杯”-BYG王者荣耀联赛数据统计报告</title>
        <style>
            * {{
                margin: 0;
                padding: 0;
                box-sizing: border-box;
            }}
            
            body {{
                font-family: 'Microsoft YaHei', Arial, sans-serif;
                font-size: 11px;
                line-height: 1.2;
                color: #333;
                background-color: #f5f7fa;
                padding: 6px;
            }}
            
            .container {{
                max-width: 100%;
                margin: 0 auto;
                background: white;
                border-radius: 3px;
                box-shadow: 0 1px 2px rgba(0,0,0,0.1);
                overflow: hidden;
            }}
            
            /* 头部样式 */
            .header {{
                background: linear-gradient(135deg, #4a6fa5 0%, #2c3e50 100%);
                color: white;
                padding: 10px 12px;
                text-align: center;
            }}
            
            .header h1 {{
                font-size: 14px;
                margin-bottom: 3px;
                font-weight: bold;
            }}
            
            .header p {{
                font-size: 9px;
                opacity: 0.9;
            }}
            
            /* 统计卡片 */
            .stats-grid {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(110px, 1fr));
                gap: 4px;
                padding: 8px;
                background: #f8f9fa;
            }}
            
            .stat-card {{
                background: white;
                border-radius: 2px;
                padding: 6px;
                text-align: center;
                box-shadow: 0 1px 1px rgba(0,0,0,0.05);
                border-left: 2px solid #4a6fa5;
            }}
            
            .stat-card .value {{
                font-size: 12px;
                font-weight: bold;
                color: #2c3e50;
                margin: 1px 0;
            }}
            
            .stat-card .label {{
                font-size: 8px;
                color: #666;
                text-transform: uppercase;
            }}
            
            /* 主要内容区 */
            .content {{
                padding: 8px;
            }}
            
            .section {{
                background: white;
                border-radius: 3px;
                padding: 8px;
                margin-bottom: 8px;
                box-shadow: 0 1px 2px rgba(0,0,0,0.05);
                border: 1px solid #eaeaea;
                page-break-inside: avoid;
            }}
            
            .section-title {{
                font-size: 11px;
                font-weight: bold;
                color: #2c3e50;
                margin-bottom: 6px;
                padding-bottom: 3px;
                border-bottom: 1px solid #eaeaea;
                display: flex;
                justify-content: space-between;
                align-items: center;
            }}
            
            .section-title span {{
                font-size: 9px;
                font-weight: normal;
                color: #666;
            }}
            
            /* 表格样式 */
            .data-table {{
                width: 100%;
                border-collapse: collapse;
                font-size: 9px;
                margin-bottom: 6px;
            }}
            
            .data-table th {{
                background-color: #f8f9fa;
                color: #4a6fa5;
                text-align: left;
                padding: 4px 6px;
                font-weight: bold;
                border-bottom: 1px solid #eaeaea;
                font-size: 8px;
                text-transform: uppercase;
            }}
            
            .data-table td {{
                padding: 3px 6px;
                border-bottom: 1px solid #f0f0f0;
            }}
            
            .data-table tr:hover {{
                background-color: #f8f9fa;
            }}
            
            .data-table .rank-1 {{ color: #e74c3c; font-weight: bold; }}
            .data-table .rank-2 {{ color: #e67e22; font-weight: bold; }}
            .data-table .rank-3 {{ color: #f1c40f; font-weight: bold; }}
            
            /* 图表容器 */
            .charts-container {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
                gap: 6px;
                margin: 6px 0;
            }}
            
            .chart-box {{
                background: white;
                border-radius: 3px;
                padding: 6px;
                box-shadow: 0 1px 1px rgba(0,0,0,0.05);
                border: 1px solid #eaeaea;
                text-align: center;
            }}
            
            .chart-box img {{
                max-width: 100%;
                height: auto;
                border-radius: 2px;
            }}
            
            .chart-title {{
                font-size: 9px;
                font-weight: bold;
                color: #2c3e50;
                margin-bottom: 4px;
            }}
            
            /* 排行榜容器 */
            .leaderboard-container {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
                gap: 6px;
                margin-top: 6px;
            }}
            
            .leaderboard-box {{
                background: #f8f9fa;
                border-radius: 3px;
                padding: 6px;
                border: 1px solid #eaeaea;
            }}
            
            .leaderboard-title {{
                font-size: 10px;
                font-weight: bold;
                color: #2c3e50;
                margin-bottom: 4px;
                text-align: center;
            }}
            
            /* 玩家卡片 */
            .player-cards {{
                display: grid;
                grid-template-columns: repeat(auto-fill, minmax(130px, 1fr));
                gap: 4px;
                margin-top: 6px;
            }}
            
            .player-card {{
                background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
                border-radius: 2px;
                padding: 5px;
                border-left: 2px solid #4a6fa5;
            }}
            
            .player-card .name {{
                font-size: 10px;
                font-weight: bold;
                color: #2c3e50;
                margin-bottom: 3px;
            }}
            
            .player-card .stats {{
                display: flex;
                justify-content: space-between;
                font-size: 8px;
                color: #666;
                margin-bottom: 1px;
            }}
            
            /* 英雄卡片 */
            .hero-cards {{
                display: grid;
                grid-template-columns: repeat(auto-fill, minmax(110px, 1fr));
                gap: 4px;
                margin-top: 6px;
            }}
            
            .hero-card {{
                background: #f8f9fa;
                border-radius: 2px;
                padding: 5px;
                text-align: center;
                border-top: 2px solid #4a6fa5;
            }}
            
            .hero-card .name {{
                font-size: 10px;
                font-weight: bold;
                color: #2c3e50;
                margin-bottom: 2px;
            }}
            
            .hero-card .win-rate {{
                font-size: 9px;
                color: #27ae60;
                font-weight: bold;
            }}
            
            /* 胜率卡片（用于英雄-玩家和玩家-英雄排行榜） */
            .winrate-cards {{
                display: grid;
                grid-template-columns: repeat(auto-fill, minmax(150px, 1fr));
                gap: 4px;
                margin-top: 6px;
            }}
            
            .winrate-card {{
                background: white;
                border-radius: 2px;
                padding: 5px;
                border: 1px solid #eaeaea;
            }}
            
            .winrate-card .title {{
                font-size: 10px;
                font-weight: bold;
                color: #2c3e50;
                margin-bottom: 3px;
                padding-bottom: 2px;
                border-bottom: 1px solid #f0f0f0;
            }}
            
            .winrate-card .item {{
                display: flex;
                justify-content: space-between;
                font-size: 8px;
                padding: 2px 0;
                border-bottom: 1px dashed #f0f0f0;
            }}
            
            .winrate-card .item:last-child {{
                border-bottom: none;
            }}
            
            .winrate-card .item .name {{
                color: #333;
            }}
            
            .winrate-card .item .rate {{
                color: #27ae60;
                font-weight: bold;
            }}
            
            /* 分路排行榜 */
            .position-grid {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
                gap: 6px;
                margin-top: 6px;
            }}
            
            .position-box {{
                background: #f8f9fa;
                border-radius: 2px;
                padding: 6px;
                border: 1px solid #eaeaea;
            }}
            
            .position-title {{
                font-size: 10px;
                font-weight: bold;
                color: #4a6fa5;
                margin-bottom: 4px;
                text-align: center;
            }}
            
            /* 底部样式 */
            .footer {{
                text-align: center;
                padding: 6px;
                color: #666;
                font-size: 8px;
                border-top: 1px solid #eaeaea;
                background: #f8f9fa;
            }}
            
            /* 响应式调整 */
            @media (max-width: 768px) {{
                .stats-grid {{
                    grid-template-columns: repeat(2, 1fr);
                }}
                
                .charts-container {{
                    grid-template-columns: 1fr;
                }}
                
                .leaderboard-container {{
                    grid-template-columns: 1fr;
                }}
                
                .position-grid {{
                    grid-template-columns: 1fr;
                }}
            }}
            
            /* 小元素样式 */
            .badge {{
                display: inline-block;
                padding: 1px 3px;
                font-size: 7px;
                border-radius: 1px;
                margin-left: 3px;
            }}
            
            .badge-blue {{ background: #4a6fa5; color: white; }}
            .badge-red {{ background: #e74c3c; color: white; }}
            .badge-green {{ background: #27ae60; color: white; }}
            .badge-gold {{ background: #f1c40f; color: #333; }}
            
            /* 紧凑列表 */
            .compact-list {{
                list-style: none;
            }}
            
            .compact-list li {{
                padding: 3px 0;
                border-bottom: 1px solid #f0f0f0;
                display: flex;
                justify-content: space-between;
                font-size: 9px;
            }}
            
            .compact-list li:last-child {{
                border-bottom: none;
            }}
            
            /* 打印样式 */
            @media print {{
                body {{
                    padding: 0;
                    font-size: 9px;
                }}
                
                .container {{
                    box-shadow: none;
                    border-radius: 0;
                }}
                
                .section {{
                    page-break-inside: avoid;
                    break-inside: avoid;
                }}
                
                .charts-container {{
                    grid-template-columns: repeat(2, 1fr);
                }}
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <!-- 头部 -->
            <div class="header">
                <h1>🏆 “饼干杯”-BYG王者荣耀联赛数据统计报告</h1>
                <p>数据统计时间: {current_time} | 共 {total_games} 场比赛 | {total_players} 名玩家 | {total_heroes} 个英雄</p>
            </div>
            
            <!-- 关键统计 -->
            <div class="section">
                <div class="section-title">📈 关键统计数据</div>
                <div class="stats-grid">
                <div class="stat-card">
                        <div class="value">{total_days}</div>
                        <div class="label">总比赛天数</div>
                    </div>
                    <div class="stat-card">
                        <div class="value">{total_games}</div>
                        <div class="label">总比赛场次</div>
                    </div>
                    <div class="stat-card">
                        <div class="value">{total_players}</div>
                        <div class="label">参赛玩家</div>
                    </div>
                    <div class="stat-card">
                        <div class="value">{total_heroes}</div>
                        <div class="label">使用英雄</div>
                    </div>
                    <div class="stat-card">
                        <div class="value">{blue_wins}</div>
                        <div class="label">蓝方胜场</div>
                    </div>
                    <div class="stat-card">
                        <div class="value">{red_wins}</div>
                        <div class="label">红方胜场</div>
                    </div>
                    <div class="stat-card">
                        <div class="value">{blue_wins/total_games*100:.1f}%</div>
                        <div class="label">蓝方胜率</div>
                    </div>
                    <div class="stat-card">
                        <div class="value">{red_wins/total_games*100:.1f}%</div>
                        <div class="label">红方胜率</div>
                    </div>
                </div>
            </div>
            
            <!-- 可视化图表 -->
            <div class="section">
                <div class="section-title">📊 数据可视化图表</div>
                <div class="charts-container">
    """
    
    # 添加图表
    if charts_dir and os.path.exists(charts_dir):
        # 选择最重要的图表
        important_charts = [
            ('胜率TOP10玩家.png', '玩家胜率TOP10'),
            ('英雄池数量TOP10.png', '英雄池TOP10'),
            ('MVP次数分布.png', 'MVP排行榜'),
            ('英雄胜率TOP10.png', '英雄胜率TOP10'),
            ('分路多样性分布.png', '分路多样性'),
            ('红蓝方胜率对比.png', '红蓝方胜率对比')
        ]
        
        for chart_file, chart_title in important_charts:
            chart_path = os.path.join(charts_dir, chart_file)
            if os.path.exists(chart_path):
                try:
                    with open(chart_path, 'rb') as img_file:
                        img_data = base64.b64encode(img_file.read()).decode('utf-8')
                    html_content += f"""
                        <div class="chart-box">
                            <div class="chart-title">{chart_title}</div>
                            <img src="data:image/png;base64,{img_data}" alt="{chart_title}">
                        </div>
                    """
                except:
                    pass
    
    html_content += """
                </div>
            </div>
            
            <!-- 玩家综合排行榜 -->
            <div class="section">
                <div class="section-title">👑 玩家综合排行榜</div>
                <table class="data-table">
                    <thead>
                        <tr>
                            <th width="40">排名</th>
                            <th>玩家</th>
                            <th width="60">胜率</th>
                            <th width="50">场次</th>
                            <th width="40">胜场</th>
                            <th width="40">MVP</th>
                            <th width="60">英雄池</th>
                        </tr>
                    </thead>
                    <tbody>
    """
    
    for idx, row in top_players.iterrows():
        rank_class = f"rank-{idx}" if idx <= 3 else ""
        html_content += f"""
                        <tr class="{rank_class}">
                            <td>{idx}</td>
                            <td>{row['玩家']}</td>
                            <td>{row['总胜率百分比']}</td>
                            <td>{row['总场次']}</td>
                            <td>{row['总胜场']}</td>
                            <td>{row['MVP次数']}</td>
                            <td>{row['英雄池数量']}</td>
                        </tr>
        """
    
    html_content += """
                    </tbody>
                </table>
            </div>
            
            <!-- 英雄排行榜 -->
            <div class="section">
                <div class="section-title">⚔️ 英雄胜率排行榜</div>
                <table class="data-table">
                    <thead>
                        <tr>
                            <th width="40">排名</th>
                            <th>英雄</th>
                            <th width="60">胜率</th>
                            <th width="50">场次</th>
                            <th width="50">胜场</th>
                        </tr>
                    </thead>
                    <tbody>
    """
    
    for idx, row in top_heroes.iterrows():
        rank_class = f"rank-{idx}" if idx <= 3 else ""
        html_content += f"""
                        <tr class="{rank_class}">
                            <td>{idx}</td>
                            <td>{row['英雄']}</td>
                            <td>{row['总胜率百分比']}</td>
                            <td>{row['总场次']}</td>
                            <td>{row['总胜场']}</td>
                        </tr>
        """
    
    html_content += """
                    </tbody>
                </table>
            </div>
            
            <!-- 同一个英雄，玩家胜率榜 -->
            <div class="section">
                <div class="section-title">🎯 同一个英雄，玩家胜率榜</div>
                <div class="winrate-cards">
    """
    
    # 选择出场次数最多的前8个英雄
    hero_games = [(hero, hero_stats[hero]['总场次']) for hero in hero_stats.keys()]
    hero_games.sort(key=lambda x: x[1], reverse=True)
    
    for hero, games in hero_games[:8]:
        if hero in hero_player_leaderboard and hero_player_leaderboard[hero]:
            html_content += f"""
                    <div class="winrate-card">
                        <div class="title">{hero} <span style="font-size:8px;color:#666;">({games}场)</span></div>
            """
            
            for i, player_data in enumerate(hero_player_leaderboard[hero][:3], 1):
                html_content += f"""
                        <div class="item">
                            <span class="name">{i}. {player_data['玩家']}</span>
                            <span class="rate">{player_data['胜率百分比']}</span>
                        </div>
                """
            
            html_content += """
                    </div>
            """
    
    html_content += """
                </div>
            </div>
            
            <!-- 同一个玩家，英雄胜率榜 -->
            <div class="section">
                <div class="section-title">🌟 同一个玩家，英雄胜率榜</div>
                <div class="winrate-cards">
    """
    
    # 选择出场次数最多的前8个玩家
    player_games = [(player, player_stats[player]['总场次']) for player in player_stats.keys()]
    player_games.sort(key=lambda x: x[1], reverse=True)
    
    for player, games in player_games[:8]:
        if player in player_hero_leaderboard and player_hero_leaderboard[player]:
            html_content += f"""
                    <div class="winrate-card">
                        <div class="title">{player} <span style="font-size:8px;color:#666;">({games}场)</span></div>
            """
            
            for i, hero_data in enumerate(player_hero_leaderboard[player][:3], 1):
                html_content += f"""
                        <div class="item">
                            <span class="name">{i}. {hero_data['英雄']}</span>
                            <span class="rate">{hero_data['胜率百分比']}</span>
                        </div>
                """
            
            html_content += """
                    </div>
            """
    
    html_content += """
                </div>
            </div>
            
            <!-- MVP排行榜和英雄池排行榜 -->
            <div class="section">
                <div class="section-title">⭐ 其他关键排行榜</div>
                <div class="leaderboard-container">
                    <!-- MVP排行榜 -->
                    <div class="leaderboard-box">
                        <div class="leaderboard-title">MVP排行榜</div>
                        <table class="data-table">
                            <thead>
                                <tr>
                                    <th width="30">排名</th>
                                    <th>玩家</th>
                                    <th width="40">次数</th>
                                </tr>
                            </thead>
                            <tbody>
    """
    
    for idx, row in top_mvp.iterrows():
        rank_class = f"rank-{idx}" if idx <= 3 else ""
        html_content += f"""
                                <tr class="{rank_class}">
                                    <td>{idx}</td>
                                    <td>{row['玩家']}</td>
                                    <td>{row['MVP次数']}</td>
                                </tr>
        """
    
    html_content += """
                            </tbody>
                        </table>
                    </div>
                    
                    <!-- 英雄池排行榜 -->
                    <div class="leaderboard-box">
                        <div class="leaderboard-title">英雄池排行榜</div>
                        <table class="data-table">
                            <thead>
                                <tr>
                                    <th width="30">排名</th>
                                    <th>玩家</th>
                                    <th width="40">数量</th>
                                    <th width="50">平均场次</th>
                                </tr>
                            </thead>
                            <tbody>
    """
    
    for idx, row in top_hero_pool.iterrows():
        rank_class = f"rank-{idx}" if idx <= 3 else ""
        avg_games = row['总场次'] / row['英雄池数量'] if row['英雄池数量'] > 0 else 0
        html_content += f"""
                                <tr class="{rank_class}">
                                    <td>{idx}</td>
                                    <td>{row['玩家']}</td>
                                    <td>{row['英雄池数量']}</td>
                                    <td>{avg_games:.1f}</td>
                                </tr>
        """
    
    html_content += """
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
            
            <!-- 分路排行榜 -->
            <div class="section">
                <div class="section-title">🎮 各分路排行榜</div>
                <div class="position-grid">
    """
    
    # 各分路排行榜
    for position in positions:
        pos_player_df = position_leaderboards[position]['player']
        pos_hero_df = position_leaderboards[position]['hero']
        
        html_content += f"""
                    <div class="position-box">
                        <div class="position-title">{position}</div>
                        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 4px;">
                            <div>
                                <div style="font-size:8px;color:#4a6fa5;margin-bottom:3px;font-weight:bold;">玩家胜率TOP5</div>
        """
        
        # 玩家TOP5
        for i, (_, row) in enumerate(pos_player_df.head(5).iterrows(), 1):
            html_content += f"""
                                <div style="font-size:9px;padding:2px 0;border-bottom:1px solid #f0f0f0;">
                                    <span>{i}. {row['玩家']}</span>
                                    <span style="float:right;color:#27ae60;">{row['胜率百分比']}</span>
                                </div>
            """
        
        html_content += """
                            </div>
                            <div>
                                <div style="font-size:8px;color:#4a6fa5;margin-bottom:3px;font-weight:bold;">英雄胜率TOP5</div>
        """
        
        # 英雄TOP5
        for i, (_, row) in enumerate(pos_hero_df.head(5).iterrows(), 1):
            html_content += f"""
                                <div style="font-size:9px;padding:2px 0;border-bottom:1px solid #f0f0f0;">
                                    <span>{i}. {row['英雄']}</span>
                                    <span style="float:right;color:#27ae60;">{row['胜率百分比']}</span>
                                </div>
            """
        
        html_content += """
                            </div>
                        </div>
                    </div>
        """
    
    html_content += """
                </div>
            </div>
            
            <!-- 玩家分路多样性分析 -->
            <div class="section">
                <div class="section-title">🔄 玩家分路多样性分析</div>
                <table class="data-table">
                    <thead>
                        <tr>
                            <th width="30">排名</th>
                            <th>玩家</th>
                            <th width="40">总场次</th>
                            <th width="50">使用分路数</th>
                            <th width="80">主要分路</th>
                            <th width="60">专注度</th>
                            <th width="50">英雄池</th>
                        </tr>
                    </thead>
                    <tbody>
    """
    
    top_diversity = position_diversity.head(10)
    for idx, row in top_diversity.iterrows():
        html_content += f"""
                        <tr>
                            <td>{idx}</td>
                            <td>{row['玩家']}</td>
                            <td>{row['总场次']}</td>
                            <td>{row['使用分路数']}</td>
                            <td>{row['主要分路']}</td>
                            <td>{row['分路专注度百分比']}</td>
                            <td>{row['英雄池数量']}</td>
                        </tr>
        """
    
    html_content += """
                    </tbody>
                </table>
            </div>
            
            <!-- 玩家详细数据（前5名） -->
            <div class="section">
                <div class="section-title">📋 玩家详细数据（TOP5）</div>
                <div class="player-cards">
    """
    
    # 添加玩家卡片
    for idx, row in top_players.head(5).iterrows():
        player = row['玩家']
        stats = player_stats[player]
        hero_pool = len(stats['英雄池'])
        
        # 统计各分路场次
        position_stats = []
        for position in positions:
            if stats[f'{position}场次'] > 0:
                position_stats.append(f"{position[:2]}:{stats[f'{position}场次']}")
        
        html_content += f"""
                    <div class="player-card">
                        <div class="name">#{idx} {player}</div>
                        <div class="stats">
                            <span>胜率: {row['总胜率百分比']}</span>
                            <span>场次: {row['总场次']}</span>
                        </div>
                        <div class="stats">
                            <span>MVP: {row['MVP次数']}</span>
                            <span>英雄池: {hero_pool}</span>
                        </div>
                        <div class="stats">
                            <span>分路: {', '.join(position_stats)}</span>
                        </div>
                    </div>
        """
    
    html_content += """
                </div>
            </div>
            
            <!-- 英雄详细数据（前5名） -->
            <div class="section">
                <div class="section-title">⚔️ 英雄详细数据（TOP5）</div>
                <div class="hero-cards">
    """
    
    # 添加英雄卡片
    for idx, row in top_heroes.head(5).iterrows():
        hero = row['英雄']
        stats = hero_stats[hero]
        
        # 计算主要分路
        main_position = ""
        max_games = 0
        for position in positions:
            if stats[f'{position}场次'] > max_games:
                max_games = stats[f'{position}场次']
                main_position = position
        
        html_content += f"""
                    <div class="hero-card">
                        <div class="name">{hero}</div>
                        <div class="win-rate">{row['总胜率百分比']}</div>
                        <div style="font-size:8px;color:#666;">出场: {row['总场次']}次</div>
                        <div style="font-size:7px;color:#999;">主要分路: {main_position}</div>
                    </div>
        """
    
    html_content += f"""
                </div>
            </div>
            
            <!-- 底部信息 -->
            <div class="footer">
                <p>报告生成时间: {current_time} | 数据来源: {total_games}场比赛 | 统计玩家: {total_players}人 | 统计英雄: {total_heroes}个</p>
                <p>© 2025 王者荣耀比赛统计系统 | Copyright: Yuanhang Zhang -- v1.0</p>
            </div>
        </div>
    </body>
    </html>
    """
    
    # 保存HTML文件
    with open('王者荣耀比赛统计报告_单页版.html', 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print("单页版HTML报告已生成: 王者荣耀比赛统计报告_单页版.html")

# 生成Excel数据文件
def generate_excel_data():
    """生成包含所有数据的Excel文件"""
    print("正在生成Excel数据文件...")
    
    with pd.ExcelWriter('王者荣耀比赛统计数据.xlsx', engine='openpyxl') as writer:
        # 原始数据
        df.to_excel(writer, sheet_name='原始比赛数据', index=False)
        
        # 核心排行榜
        player_leaderboard.to_excel(writer, sheet_name='玩家综合排行榜')
        hero_leaderboard.to_excel(writer, sheet_name='英雄胜率排行榜')
        mvp_leaderboard.to_excel(writer, sheet_name='MVP排行榜')
        hero_pool_leaderboard.to_excel(writer, sheet_name='英雄池排行榜')
        position_diversity.to_excel(writer, sheet_name='分路多样性')
        
        # 各分路排行榜
        for position in positions:
            pos_player_df = position_leaderboards[position]['player']
            pos_hero_df = position_leaderboards[position]['hero']
            
            pos_player_df.to_excel(writer, sheet_name=f'{position}玩家榜')
            pos_hero_df.to_excel(writer, sheet_name=f'{position}英雄榜')
        
        # 玩家详细数据
        player_detail_data = []
        for player, stats in player_stats.items():
            if stats['总场次'] > 0:
                player_detail_data.append({
                    '玩家': player,
                    '总场次': stats['总场次'],
                    '总胜场': stats['总胜场'],
                    '总胜率': stats['总胜场'] / stats['总场次'],
                    'MVP次数': stats['MVP次数'],
                    '英雄池数量': len(stats['英雄池']),
                    '英雄池列表': ', '.join(sorted(stats['英雄池'])),
                    '使用分路数': sum(1 for pos in positions if stats[f'{pos}场次'] > 0),
                    '主要分路': max(positions, 
                                  key=lambda pos: stats[f'{pos}场次']) if stats['总场次'] > 0 else '无'
                })
        
        player_detail_df = pd.DataFrame(player_detail_data)
        player_detail_df.to_excel(writer, sheet_name='玩家详细数据', index=False)
        
        # 英雄详细数据
        hero_detail_data = []
        for hero, stats in hero_stats.items():
            if stats['总场次'] > 0:
                hero_detail_data.append({
                    '英雄': hero,
                    '总场次': stats['总场次'],
                    '总胜场': stats['总胜场'],
                    '总胜率': stats['总胜场'] / stats['总场次'],
                    '边路场次': stats['边路场次'],
                    '打野场次': stats['打野场次'],
                    '中路场次': stats['中路场次'],
                    '发育路场次': stats['发育路场次'],
                    '游走场次': stats['游走场次']
                })
        
        hero_detail_df = pd.DataFrame(hero_detail_data)
        hero_detail_df.to_excel(writer, sheet_name='英雄详细数据', index=False)
        
        # 同一个英雄，玩家胜率榜
        hero_player_data = []
        for hero, player_list in hero_player_leaderboard.items():
            for player_data in player_list:
                hero_player_data.append({
                    '英雄': hero,
                    '玩家': player_data['玩家'],
                    '场次': player_data['场次'],
                    '胜场': player_data['胜场'],
                    '胜率': player_data['胜率']
                })
        
        if hero_player_data:
            hero_player_df = pd.DataFrame(hero_player_data)
            hero_player_df.to_excel(writer, sheet_name='英雄玩家胜率榜', index=False)
        
        # 同一个玩家，英雄胜率榜
        player_hero_data = []
        for player, hero_list in player_hero_leaderboard.items():
            for hero_data in hero_list:
                player_hero_data.append({
                    '玩家': player,
                    '英雄': hero_data['英雄'],
                    '场次': hero_data['场次'],
                    '胜场': hero_data['胜场'],
                    '胜率': hero_data['胜率']
                })
        
        if player_hero_data:
            player_hero_df = pd.DataFrame(player_hero_data)
            player_hero_df.to_excel(writer, sheet_name='玩家英雄胜率榜', index=False)
    
    print("Excel数据文件已生成: 王者荣耀比赛统计数据.xlsx")

# 生成PDF报告（可选）
def generate_pdf_report():
    """生成PDF报告（需要weasyprint）"""
    try:
        from weasyprint import HTML
        
        print("正在生成PDF报告...")
        
        # 读取HTML内容
        with open('王者荣耀比赛统计报告_单页版.html', 'r', encoding='utf-8') as f:
            html_content = f.read()
        
        # 生成PDF
        HTML(string=html_content).write_pdf('王者荣耀比赛统计报告.pdf')
        
        print("PDF报告已生成: 王者荣耀比赛统计报告.pdf")
        
    except ImportError:
        print("警告: weasyprint库未安装，无法生成PDF报告")
        print("请使用以下命令安装: pip install weasyprint")
    except Exception as e:
        print(f"PDF生成错误: {e}")

# 主程序
def main():
    print("\n" + "="*80)
    print("王者荣耀比赛统计报表生成系统")
    print("="*80)
    
    # 生成所有报告
    generate_html_report()
    generate_excel_data()
    generate_pdf_report()
    
    print("\n" + "="*80)
    print("报告生成完成！")
    print("="*80)
    print("\n生成的文件:")
    print("1. 王者荣耀比赛统计报告_单页版.html - 单页完整HTML报告")
    print("2. 王者荣耀比赛统计数据.xlsx - 所有数据的Excel文件")
    print("3. 王者荣耀比赛统计报告.pdf - PDF格式报告（如果已安装weasyprint）")
    print("4. charts/ - 包含所有可视化图表的目录")
    print("\n打开 王者荣耀比赛统计报告_单页版.html 查看完整报告")

# 运行主程序
if __name__ == "__main__":
    main()