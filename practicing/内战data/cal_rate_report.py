import pandas as pd
import numpy as np
from collections import defaultdict
import warnings
import os
from datetime import datetime
import base64
from itertools import combinations

warnings.filterwarnings('ignore')

# 创建示例数据
data = pd.read_excel('C:/Files/Ubiquant/code/HOK/hok_bp/practicing/内战data/内战计分表 - 2026.xlsx')
data['比赛时间'] = data['比赛时间'].apply(lambda x: x.strftime("%Y-%m-%d"))

# 创建列名
columns = ['比赛ID', '比赛时间', '胜方', '蓝方边', '蓝方野', '蓝方中', '蓝方射', '蓝方辅', '蓝方MVP', 
           '红方边', '红方野', '红方中', '红方射', '红方辅', '红方MVP']

# 创建DataFrame
df = data[columns].copy()

# ========== 按年份分组数据 ==========
df_2025 = df[df['比赛时间'].str.startswith('2025')].copy()
df_2026 = df[df['比赛时间'].str.startswith('2026')].copy()
df_all = df.copy()
df = df_2026

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
    '英雄胜场': defaultdict(int),
    '英雄场次': defaultdict(int),
    '队友配合': defaultdict(int),  # 新增：统计与不同队友的合作次数
    '对手交手': defaultdict(int),  # 新增：统计与不同对手的交手次数
    '每日比赛': defaultdict(int),  # 新增：统计每日比赛场次
    '连胜': 0,  # 新增：当前连胜次数
    '最长连胜': 0,  # 新增：最长连胜次数
})

hero_stats = defaultdict(lambda: {
    '总场次': 0, '总胜场': 0,
    '边路场次': 0, '边路胜场': 0,
    '打野场次': 0, '打野胜场': 0,
    '中路场次': 0, '中路胜场': 0,
    '发育路场次': 0, '发育路胜场': 0,
    '游走场次': 0, '游走胜场': 0,
    '玩家胜场': defaultdict(int),
    '玩家场次': defaultdict(int),
    '最佳搭档': defaultdict(int),  # 新增：英雄搭配次数
    '克制关系': defaultdict(int),  # 新增：英雄克制关系
})

# 新增统计：玩家组合数据
player_pair_stats = defaultdict(lambda: {
    '一起出场': 0,
    '一起胜利': 0,
    '一起失败': 0,
})

# 新增统计：英雄组合数据
hero_pair_stats = defaultdict(lambda: {
    '一起出场': 0,
    '一起胜利': 0,
    '一起失败': 0,
})

# 处理每场比赛
match_history = []  # 记录每场比赛的详细信息，用于连胜统计

for idx, row in df.iterrows():
    winner = row['胜方']
    match_id = row['比赛ID']
    match_date = row['比赛时间']
    
    # 记录比赛信息
    match_info = {
        'id': match_id,
        'date': match_date,
        'winner': winner,
        'blue_players': [],
        'red_players': [],
        'blue_heroes': [],
        'red_heroes': []
    }
    
    # 处理蓝方
    blue_players = []
    for pos_col, position in position_map.items():
        if pos_col.startswith('蓝方'):
            cell_value = row[pos_col]
            if isinstance(cell_value, str) and '-' in cell_value:
                player, hero = cell_value.split('-', 1)
                blue_players.append(player)
                match_info['blue_players'].append(player)
                match_info['blue_heroes'].append(hero)
                
                # 更新玩家统计
                player_stats[player]['总场次'] += 1
                player_stats[player][f'{position}场次'] += 1
                player_stats[player]['每日比赛'][match_date] += 1
                
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
    red_players = []
    for pos_col, position in position_map.items():
        if pos_col.startswith('红方'):
            cell_value = row[pos_col]
            if isinstance(cell_value, str) and '-' in cell_value:
                player, hero = cell_value.split('-', 1)
                red_players.append(player)
                match_info['red_players'].append(player)
                match_info['red_heroes'].append(hero)
                
                # 更新玩家统计
                player_stats[player]['总场次'] += 1
                player_stats[player][f'{position}场次'] += 1
                player_stats[player]['每日比赛'][match_date] += 1
                
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
    
    # 更新玩家配合数据（同一方玩家）
    for player1, player2 in combinations(blue_players, 2):
        key = tuple(sorted([player1, player2]))
        player_stats[player1]['队友配合'][player2] += 1
        player_stats[player2]['队友配合'][player1] += 1
        player_pair_stats[key]['一起出场'] += 1
        if winner == '蓝':
            player_pair_stats[key]['一起胜利'] += 1
        else:
            player_pair_stats[key]['一起失败'] += 1
    
    for player1, player2 in combinations(red_players, 2):
        key = tuple(sorted([player1, player2]))
        player_stats[player1]['队友配合'][player2] += 1
        player_stats[player2]['队友配合'][player1] += 1
        player_pair_stats[key]['一起出场'] += 1
        if winner == '红':
            player_pair_stats[key]['一起胜利'] += 1
        else:
            player_pair_stats[key]['一起失败'] += 1
    
    # 更新玩家交手数据（不同方玩家）
    for blue_player in blue_players:
        for red_player in red_players:
            player_stats[blue_player]['对手交手'][red_player] += 1
            player_stats[red_player]['对手交手'][blue_player] += 1
    
    # 更新英雄配合数据（同一方英雄）
    for hero1, hero2 in combinations(match_info['blue_heroes'], 2):
        key = tuple(sorted([hero1, hero2]))
        hero_pair_stats[key]['一起出场'] += 1
        if winner == '蓝':
            hero_pair_stats[key]['一起胜利'] += 1
        else:
            hero_pair_stats[key]['一起失败'] += 1
    
    for hero1, hero2 in combinations(match_info['red_heroes'], 2):
        key = tuple(sorted([hero1, hero2]))
        hero_pair_stats[key]['一起出场'] += 1
        if winner == '红':
            hero_pair_stats[key]['一起胜利'] += 1
        else:
            hero_pair_stats[key]['一起失败'] += 1
    
    # 更新英雄克制数据（不同方英雄）
    for blue_hero in match_info['blue_heroes']:
        for red_hero in match_info['red_heroes']:
            hero_stats[blue_hero]['克制关系'][red_hero] += 1
    
    match_history.append(match_info)

# 计算玩家连胜数据
player_last_result = {}
for player in player_stats.keys():
    player_last_result[player] = None  # None表示未知，'W'表示胜利，'L'表示失败

# 按照比赛时间排序
df_sorted = df.sort_values('比赛时间')
for idx, row in df_sorted.iterrows():
    winner = row['胜方']
    match_date = row['比赛时间']
    
    # 获取蓝方玩家
    blue_players = []
    for pos_col in position_map.keys():
        if pos_col.startswith('蓝方'):
            cell_value = row[pos_col]
            if isinstance(cell_value, str) and '-' in cell_value:
                player, _ = cell_value.split('-', 1)
                blue_players.append(player)
    
    # 获取红方玩家
    red_players = []
    for pos_col in position_map.keys():
        if pos_col.startswith('红方'):
            cell_value = row[pos_col]
            if isinstance(cell_value, str) and '-' in cell_value:
                player, _ = cell_value.split('-', 1)
                red_players.append(player)
    
    # 更新每个玩家的连胜数据
    for player in blue_players:
        if winner == '蓝':
            player_stats[player]['连胜'] += 1
            player_last_result[player] = 'W'
        else:
            player_stats[player]['连胜'] = 0
            player_last_result[player] = 'L'
        
        if player_stats[player]['连胜'] > player_stats[player]['最长连胜']:
            player_stats[player]['最长连胜'] = player_stats[player]['连胜']
    
    for player in red_players:
        if winner == '红':
            player_stats[player]['连胜'] += 1
            player_last_result[player] = 'W'
        else:
            player_stats[player]['连胜'] = 0
            player_last_result[player] = 'L'
        
        if player_stats[player]['连胜'] > player_stats[player]['最长连胜']:
            player_stats[player]['最长连胜'] = player_stats[player]['连胜']

# ========== 赏金榜计算（仅2026年） ==========
def calculate_bounty_leaderboard(df_year):
    """
    计算赏金榜（增强版）
    规则：胜方全员+2元，胜方MVP+1元，败方MVP+1元
    返回: (排行榜DataFrame, 奖金池信息dict, 日期列表)
    """
    # 2026年奖金池初始值
    POOL_INITIAL = 401.5
    
    # 总赏金统计
    bounty_total = defaultdict(float)
    # 每日赏金明细 {日期: {玩家: 金额}}
    bounty_daily = defaultdict(lambda: defaultdict(float))
    # 获取所有日期
    all_dates = sorted(df_year['比赛时间'].unique().tolist())
    
    for idx, row in df_year.iterrows():
        winner = row['胜方']
        match_date = row['比赛时间']
        
        # 蓝方玩家
        blue_players = []
        for pos_col in ['蓝方边', '蓝方野', '蓝方中', '蓝方射', '蓝方辅']:
            cell_value = row[pos_col]
            if isinstance(cell_value, str) and '-' in cell_value:
                player, _ = cell_value.split('-', 1)
                blue_players.append(player)
        # 红方玩家
        red_players = []
        for pos_col in ['红方边', '红方野', '红方中', '红方射', '红方辅']:
            cell_value = row[pos_col]
            if isinstance(cell_value, str) and '-' in cell_value:
                player, _ = cell_value.split('-', 1)
                red_players.append(player)
        
        # 胜方全员+2元
        win_players = blue_players if winner == '蓝' else red_players
        lose_players = red_players if winner == '蓝' else blue_players
        for p in win_players:
            bounty_total[p] += 2
            bounty_daily[match_date][p] += 2
        
        # 胜方MVP+1元
        mvp_pos = row['蓝方MVP'] if winner == '蓝' else row['红方MVP']
        mvp_map = {'边': 0, '野': 1, '中': 2, '射': 3, '辅': 4}
        try:
            mvp_idx = mvp_map.get(mvp_pos, None)
            if mvp_idx is not None and mvp_idx < len(win_players):
                bounty_total[win_players[mvp_idx]] += 1
                bounty_daily[match_date][win_players[mvp_idx]] += 1
        except:
            pass
        
        # 败方MVP+1元
        lose_mvp_pos = row['红方MVP'] if winner == '蓝' else row['蓝方MVP']
        try:
            lose_mvp_idx = mvp_map.get(lose_mvp_pos, None)
            if lose_mvp_idx is not None and lose_mvp_idx < len(lose_players):
                bounty_total[lose_players[lose_mvp_idx]] += 1
                bounty_daily[match_date][lose_players[lose_mvp_idx]] += 1
        except:
            pass
    
    # 计算总发放金额
    total_distributed = sum(bounty_total.values())
    
    # 生成排行榜数据（含每日明细列）
    leaderboard_data = []
    for player, total in bounty_total.items():
        row_data = {
            '玩家': player,
            '总赏金': total,
        }
        # 添加每日赏金列
        for date in all_dates:
            daily_amount = bounty_daily[date].get(player, 0)
            row_data[date] = daily_amount if daily_amount > 0 else 0
        leaderboard_data.append(row_data)
    
    # 按总赏金排序
    bounty_df = pd.DataFrame(leaderboard_data)
    if len(bounty_df) > 0:
        bounty_df = bounty_df.sort_values(by='总赏金', ascending=False).reset_index(drop=True)
        bounty_df.index = bounty_df.index + 1
        bounty_df.index.name = '排名'
    
    # 奖金池信息
    pool_info = {
        'initial': POOL_INITIAL,
        'distributed': total_distributed,
        'remaining': POOL_INITIAL - total_distributed
    }
    
    return bounty_df, pool_info, all_dates

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
                '英雄池数量': len(stats['英雄池']),
                '最长连胜': stats['最长连胜'],
                '当前连胜': stats['连胜']
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
                '总胜率百分比': f"{win_rate * 100:.2f}%",
                '使用玩家数': len(stats['玩家场次'])
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
                'MVP次数': stats['MVP次数'],
                '总场次': stats['总场次'],
                'MVP率': f"{stats['MVP次数']/stats['总场次']*100:.1f}%" if stats['总场次'] > 0 else "0%"
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
                '分路多样性': sum(1 for pos in ['边路', '打野', '中路', '发育路', '游走'] if stats[f'{pos}场次'] > 0)
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
                '胜率百分比': f"{win_rate * 100:.2f}%",
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
                '胜率百分比': f"{win_rate * 100:.2f}%",
                '使用玩家数': len([p for p in stats['玩家场次'] if stats['玩家场次'][p] > 0])
            })
    
    leaderboard_df = pd.DataFrame(leaderboard)
    leaderboard_df = leaderboard_df.sort_values(by=['胜率', '场次'], ascending=[False, False])
    leaderboard_df = leaderboard_df.reset_index(drop=True)
    leaderboard_df.index = leaderboard_df.index + 1
    leaderboard_df.index.name = '排名'
    
    return leaderboard_df

# 7. 同一个英雄，玩家胜率榜
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
            hero_player_stats[hero] = player_list
    
    return hero_player_stats

# 8. 同一个玩家，英雄胜率榜
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
            player_hero_stats[player] = hero_list
    
    return player_hero_stats

# 9. 最佳队友组合榜
def create_best_teammate_leaderboard():
    """生成最佳队友组合排行榜"""
    leaderboard = []
    
    for (player1, player2), stats in player_pair_stats.items():
        if stats['一起出场'] >= 2:  # 至少一起出场2次
            win_rate = stats['一起胜利'] / stats['一起出场'] if stats['一起出场'] > 0 else 0
            leaderboard.append({
                '队友组合': f"{player1} & {player2}",
                '一起出场': stats['一起出场'],
                '一起胜利': stats['一起胜利'],
                '一起失败': stats['一起失败'],
                '胜率': win_rate,
                '胜率百分比': f"{win_rate * 100:.2f}%"
            })
    
    leaderboard_df = pd.DataFrame(leaderboard)
    leaderboard_df = leaderboard_df.sort_values(by=['胜率', '一起出场'], ascending=[False, False])
    leaderboard_df = leaderboard_df.reset_index(drop=True)
    leaderboard_df.index = leaderboard_df.index + 1
    leaderboard_df.index.name = '排名'
    
    return leaderboard_df

# 10. 最佳英雄组合榜
def create_best_hero_combo_leaderboard():
    """生成最佳英雄组合排行榜"""
    leaderboard = []
    
    for (hero1, hero2), stats in hero_pair_stats.items():
        if stats['一起出场'] >= 2:  # 至少一起出场2次
            win_rate = stats['一起胜利'] / stats['一起出场'] if stats['一起出场'] > 0 else 0
            leaderboard.append({
                '英雄组合': f"{hero1} & {hero2}",
                '一起出场': stats['一起出场'],
                '一起胜利': stats['一起胜利'],
                '一起失败': stats['一起失败'],
                '胜率': win_rate,
                '胜率百分比': f"{win_rate * 100:.2f}%"
            })
    
    leaderboard_df = pd.DataFrame(leaderboard)
    leaderboard_df = leaderboard_df.sort_values(by=['胜率', '一起出场'], ascending=[False, False])
    leaderboard_df = leaderboard_df.reset_index(drop=True)
    leaderboard_df.index = leaderboard_df.index + 1
    leaderboard_df.index.name = '排名'
    
    return leaderboard_df

# 11. 每日比赛统计
def create_daily_stats():
    """生成每日比赛统计"""
    daily_stats = defaultdict(lambda: {'比赛场次': 0, '蓝方胜': 0, '红方胜': 0})
    
    for idx, row in df.iterrows():
        date = row['比赛时间']
        winner = row['胜方']
        
        daily_stats[date]['比赛场次'] += 1
        if winner == '蓝':
            daily_stats[date]['蓝方胜'] += 1
        else:
            daily_stats[date]['红方胜'] += 1
    
    leaderboard = []
    for date, stats in daily_stats.items():
        blue_rate = stats['蓝方胜'] / stats['比赛场次'] if stats['比赛场次'] > 0 else 0
        red_rate = stats['红方胜'] / stats['比赛场次'] if stats['比赛场次'] > 0 else 0
        
        leaderboard.append({
            '日期': date,
            '比赛场次': stats['比赛场次'],
            '蓝方胜场': stats['蓝方胜'],
            '红方胜场': stats['红方胜'],
            '蓝方胜率': f"{blue_rate * 100:.1f}%",
            '红方胜率': f"{red_rate * 100:.1f}%"
        })
    
    leaderboard_df = pd.DataFrame(leaderboard)
    leaderboard_df = leaderboard_df.sort_values(by='日期', ascending=True)
    leaderboard_df = leaderboard_df.reset_index(drop=True)
    leaderboard_df.index = leaderboard_df.index + 1
    leaderboard_df.index.name = '序号'
    
    return leaderboard_df

# 12. 玩家活跃度排行榜
def create_player_activity_leaderboard():
    """生成玩家活跃度排行榜（按比赛天数）"""
    leaderboard = []
    
    for player, stats in player_stats.items():
        if stats['总场次'] > 0:
            active_days = len(stats['每日比赛'])
            avg_games_per_day = stats['总场次'] / active_days if active_days > 0 else 0
            
            leaderboard.append({
                '玩家': player,
                '总场次': stats['总场次'],
                '活跃天数': active_days,
                '场均比赛': f"{avg_games_per_day:.1f}",
                '最近比赛': max(stats['每日比赛'].keys()) if stats['每日比赛'] else '无'
            })
    
    leaderboard_df = pd.DataFrame(leaderboard)
    leaderboard_df = leaderboard_df.sort_values(by=['活跃天数', '总场次'], ascending=[False, False])
    leaderboard_df = leaderboard_df.reset_index(drop=True)
    leaderboard_df.index = leaderboard_df.index + 1
    leaderboard_df.index.name = '排名'
    
    return leaderboard_df

# 13. 连胜排行榜
def create_streak_leaderboard():
    """生成连胜排行榜"""
    leaderboard = []
    
    for player, stats in player_stats.items():
        if stats['总场次'] > 0:
            leaderboard.append({
                '玩家': player,
                '最长连胜': stats['最长连胜'],
                '当前连胜': stats['连胜'],
                '总场次': stats['总场次'],
                '总胜率': f"{stats['总胜场']/stats['总场次']*100:.2f}%" if stats['总场次'] > 0 else "0%"
            })
    
    leaderboard_df = pd.DataFrame(leaderboard)
    leaderboard_df = leaderboard_df.sort_values(by=['最长连胜', '当前连胜'], ascending=[False, False])
    leaderboard_df = leaderboard_df.reset_index(drop=True)
    leaderboard_df.index = leaderboard_df.index + 1
    leaderboard_df.index.name = '排名'
    
    return leaderboard_df

# 14. 英雄克制矩阵
def create_hero_counter_matrix():
    """
    生成英雄克制矩阵
    统计每个英雄对阵其他英雄时的胜率
    返回：克制数据字典、克制排行榜DataFrame、被克制排行榜DataFrame
    """
    # 收集对战数据：{(我方英雄, 敌方英雄): {'对战': n, '胜利': m}}
    counter_data = defaultdict(lambda: {'对战': 0, '胜利': 0})
    
    for match in match_history:
        winner = match['winner']
        blue_heroes = match['blue_heroes']
        red_heroes = match['red_heroes']
        
        # 蓝方英雄 vs 红方英雄
        for blue_hero in blue_heroes:
            for red_hero in red_heroes:
                counter_data[(blue_hero, red_hero)]['对战'] += 1
                if winner == '蓝':
                    counter_data[(blue_hero, red_hero)]['胜利'] += 1
        
        # 红方英雄 vs 蓝方英雄
        for red_hero in red_heroes:
            for blue_hero in blue_heroes:
                counter_data[(red_hero, blue_hero)]['对战'] += 1
                if winner == '红':
                    counter_data[(red_hero, blue_hero)]['胜利'] += 1
    
    return counter_data

def create_hero_counter_leaderboard(counter_data, min_games=2):
    """
    生成英雄克制排行榜
    找出克制关系最明显的英雄对
    """
    leaderboard = []
    
    for (hero, enemy), stats in counter_data.items():
        if stats['对战'] >= min_games:
            win_rate = stats['胜利'] / stats['对战']
            leaderboard.append({
                '英雄': hero,
                '对手': enemy,
                '对战次数': stats['对战'],
                '胜场': stats['胜利'],
                '胜率': win_rate,
                '胜率百分比': f"{win_rate * 100:.1f}%",
                '克制程度': '🔥强克' if win_rate >= 0.7 else ('✓克制' if win_rate >= 0.6 else '均势')
            })
    
    leaderboard_df = pd.DataFrame(leaderboard)
    if len(leaderboard_df) > 0:
        leaderboard_df = leaderboard_df.sort_values(by=['胜率', '对战次数'], ascending=[False, False])
        leaderboard_df = leaderboard_df.reset_index(drop=True)
        leaderboard_df.index = leaderboard_df.index + 1
        leaderboard_df.index.name = '排名'
    
    return leaderboard_df

def create_hero_counter_summary(counter_data, min_games=1):
    """
    生成每个英雄的克制总结
    包括：最克制的对手、最被克制的对手
    """
    hero_summary = defaultdict(lambda: {
        '总对战': 0,
        '总胜利': 0,
        '克制对手': [],  # [(对手, 胜率, 场次)]
        '被克制对手': []  # [(对手, 胜率, 场次)]
    })
    
    for (hero, enemy), stats in counter_data.items():
        if stats['对战'] >= min_games:
            win_rate = stats['胜利'] / stats['对战']
            hero_summary[hero]['总对战'] += stats['对战']
            hero_summary[hero]['总胜利'] += stats['胜利']
            
            if win_rate >= 0.55:  # 克制对手
                hero_summary[hero]['克制对手'].append((enemy, win_rate, stats['对战']))
            elif win_rate <= 0.45:  # 被克制
                hero_summary[hero]['被克制对手'].append((enemy, win_rate, stats['对战']))
    
    # 排序
    for hero in hero_summary:
        hero_summary[hero]['克制对手'].sort(key=lambda x: x[1], reverse=True)
        hero_summary[hero]['被克制对手'].sort(key=lambda x: x[1])
    
    return hero_summary

def create_hero_matchup_matrix(counter_data):
    """
    生成英雄对战矩阵（用于热力图）
    返回DataFrame，行为我方英雄，列为敌方英雄，值为胜率
    """
    # 获取所有英雄
    all_heroes = set()
    for (hero, enemy) in counter_data.keys():
        all_heroes.add(hero)
        all_heroes.add(enemy)
    
    all_heroes = sorted(list(all_heroes))
    
    # 创建矩阵
    matrix_data = []
    for hero in all_heroes:
        row = {'英雄': hero}
        for enemy in all_heroes:
            if hero == enemy:
                row[enemy] = None  # 自己对自己
            elif (hero, enemy) in counter_data and counter_data[(hero, enemy)]['对战'] > 0:
                stats = counter_data[(hero, enemy)]
                row[enemy] = stats['胜利'] / stats['对战']
            else:
                row[enemy] = None
        matrix_data.append(row)
    
    matrix_df = pd.DataFrame(matrix_data)
    matrix_df = matrix_df.set_index('英雄')
    
    return matrix_df

# 生成所有排行榜
print("正在生成排行榜数据...")
player_leaderboard = create_player_leaderboard()
hero_leaderboard = create_hero_leaderboard()
mvp_leaderboard = create_mvp_leaderboard()
hero_pool_leaderboard = create_hero_pool_leaderboard()
best_teammate_leaderboard = create_best_teammate_leaderboard()
# best_hero_combo_leaderboard = create_best_hero_combo_leaderboard()
best_hero_combo_leaderboard = pd.DataFrame()

daily_stats = create_daily_stats()
player_activity_leaderboard = create_player_activity_leaderboard()
streak_leaderboard = create_streak_leaderboard()

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
    import matplotlib.font_manager as fm
    import seaborn as sns
    
    # ==================== 字体配置优化 ====================
    # 优先使用更美观的中文字体
    # Windows系统常见美观字体优先级：微软雅黑 > 思源黑体 > 苹方 > 黑体
    preferred_fonts = [
        'Microsoft YaHei',      # 微软雅黑 - 现代感强，推荐
        'Source Han Sans CN',   # 思源黑体 - 开源优质字体
        'PingFang SC',          # 苹方 - macOS系统字体
        'Noto Sans CJK SC',     # Noto思源黑体
        'WenQuanYi Micro Hei',  # 文泉驿微米黑
        'SimHei',               # 黑体 - 系统自带
        'Arial Unicode MS',     # Arial Unicode
        'DejaVu Sans'           # 后备字体
    ]
    
    # 检测可用字体
    available_fonts = [f.name for f in fm.fontManager.ttflist]
    selected_font = None
    for font in preferred_fonts:
        if font in available_fonts:
            selected_font = font
            print(f"图表使用字体: {font}")
            break
    
    if selected_font is None:
        selected_font = 'SimHei'
        print(f"未找到优选字体，使用默认字体: SimHei")
    
    # 设置字体配置
    plt.rcParams['font.sans-serif'] = [selected_font] + preferred_fonts
    plt.rcParams['font.family'] = 'sans-serif'
    plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题
    
    # 字体样式优化
    plt.rcParams['font.size'] = 10              # 基础字号
    plt.rcParams['axes.titlesize'] = 13         # 标题字号
    plt.rcParams['axes.titleweight'] = 'bold'   # 标题加粗
    plt.rcParams['axes.labelsize'] = 11         # 轴标签字号
    plt.rcParams['xtick.labelsize'] = 9         # X轴刻度字号
    plt.rcParams['ytick.labelsize'] = 9         # Y轴刻度字号
    plt.rcParams['legend.fontsize'] = 9         # 图例字号
    
    # 设置统一的柔和配色方案
    # ==================== 配色方案 ====================
    CHART_COLORS = {
        'primary': '#5B7BA3',      # 主色-柔和蓝
        'secondary': '#7BA3C9',    # 次色-浅蓝
        'accent': '#E8A87C',       # 强调色-暖橙
        'success': '#7CB798',      # 成功色-柔和绿
        'warning': '#E8C87C',      # 警告色-柔和黄
        'danger': '#C98B8B',       # 危险色-柔和红
        'bg': '#F7F9FC',           # 背景色
        'text': '#4A5568',         # 文字色
        'border': '#CBD5E0'        # 边框色
    }
    
    # 柔和渐变色板
    soft_blues = ['#8FB8DE', '#7BA3C9', '#6A91B8', '#5B7FA7', '#4C6D96', '#3D5B85', '#2E4974', '#1F3763']
    soft_warm = ['#F4D9C6', '#E8C9A8', '#E0B98A', '#D8A96C', '#D0994E', '#C88930', '#B07928', '#986920']
    soft_gradient = ['#5B7BA3', '#6A8FB3', '#79A3C3', '#88B7D3', '#97CBE3', '#A6DFF3', '#B5E3F8', '#C4E7FC']
    
    # ==================== 全局图表样式 ====================
    plt.rcParams['axes.facecolor'] = '#FAFBFC'
    plt.rcParams['figure.facecolor'] = '#FFFFFF'
    plt.rcParams['axes.edgecolor'] = CHART_COLORS['border']
    plt.rcParams['axes.labelcolor'] = CHART_COLORS['text']
    plt.rcParams['xtick.color'] = CHART_COLORS['text']
    plt.rcParams['ytick.color'] = CHART_COLORS['text']
    plt.rcParams['axes.titlecolor'] = CHART_COLORS['text']
    plt.rcParams['axes.spines.top'] = False      # 隐藏顶部边框
    plt.rcParams['axes.spines.right'] = False    # 隐藏右侧边框
    plt.rcParams['figure.dpi'] = 150             # 提高图表清晰度
    plt.rcParams['savefig.dpi'] = 150            # 保存时的DPI
    plt.rcParams['axes.titlepad'] = 12           # 标题与图表间距
    plt.rcParams['axes.labelpad'] = 8            # 轴标签与轴间距
    
    # 创建图表保存目录
    charts_dir = "charts"
    if not os.path.exists(charts_dir):
        os.makedirs(charts_dir)
    
    # 1. 胜率TOP10玩家图表
    fig1, ax1 = plt.subplots(figsize=(10, 6))
    top_players = player_leaderboard.head(10).copy()
    top_players = top_players.sort_values(by='总胜率',ascending=False)
    colors = [soft_blues[i % len(soft_blues)] for i in range(len(top_players))]
    bars = ax1.barh(top_players['玩家'], top_players['总胜率'], color=colors, edgecolor='white', linewidth=0.5)
    ax1.set_xlabel('胜率')
    ax1.set_title('玩家胜率TOP10')
    ax1.set_xlim(0, 1)
    ax1.invert_yaxis()
    
    # 添加数值标签
    for i, (bar, row) in enumerate(zip(bars, top_players.itertuples())):
        ax1.text(bar.get_width() + 0.01, bar.get_y() + bar.get_height()/2,
                f'{row.总胜率:.2%}', ha='left', va='center', fontsize=9, color=CHART_COLORS['text'])
    
    plt.tight_layout()
    plt.savefig(f'{charts_dir}/玩家胜率TOP10.png', bbox_inches='tight', facecolor='white')
    plt.close()
    
    # 2. 英雄池数量TOP10图表
    fig2, ax2 = plt.subplots(figsize=(10, 6))
    hero_pool_top10 = hero_pool_leaderboard.head(10).copy()
    hero_pool_top10 = hero_pool_top10.sort_values(by='英雄池数量',ascending=False)
    colors = [soft_warm[i % len(soft_warm)] for i in range(len(hero_pool_top10))]
    bars = ax2.barh(hero_pool_top10['玩家'], hero_pool_top10['英雄池数量'], color=colors, edgecolor='white', linewidth=0.5)
    ax2.set_xlabel('英雄池数量')
    ax2.set_title('英雄池数量TOP10玩家')
    ax2.invert_yaxis()
    
    # 添加数值标签
    for bar in bars:
        width = bar.get_width()
        ax2.text(width + 0.1, bar.get_y() + bar.get_height()/2,
                f'{int(width)}', ha='left', va='center', fontsize=9, color=CHART_COLORS['text'])
    
    plt.tight_layout()
    plt.savefig(f'{charts_dir}/英雄池数量TOP10.png', bbox_inches='tight', facecolor='white')
    plt.close()
    
    # 3. MVP次数分布图表
    fig3, ax3 = plt.subplots(figsize=(8, 5))
    mvp_data = mvp_leaderboard.head(10).copy()
    if len(mvp_data) > 0:
        mvp_data = mvp_data.sort_values(by='MVP次数',ascending=False)
        # 使用柔和的金色和蓝色
        colors = [CHART_COLORS['accent'] if i == 0 else CHART_COLORS['secondary'] for i in range(len(mvp_data))]
        bars = ax3.barh(mvp_data['玩家'], mvp_data['MVP次数'], color=colors, edgecolor='white', linewidth=0.5)
        ax3.set_xlabel('MVP次数')
        ax3.set_title('MVP次数TOP10')
        ax3.invert_yaxis()
        
        # 添加数值标签
        for bar in bars:
            width = bar.get_width()
            ax3.text(width + 0.05, bar.get_y() + bar.get_height()/2,
                    f'{int(width)}', ha='left', va='center', fontsize=9, color=CHART_COLORS['text'])
    
    plt.tight_layout()
    plt.savefig(f'{charts_dir}/MVP次数分布.png', bbox_inches='tight', facecolor='white')
    plt.close()
    
    # 4. 英雄胜率TOP10图表
    fig4, ax4 = plt.subplots(figsize=(10, 6))
    top_heroes = hero_leaderboard[hero_leaderboard['总场次'] >= 2].head(10).copy()
    if len(top_heroes) > 0:
        top_heroes = top_heroes.sort_values(by='总胜率',ascending=False)
        colors = [soft_gradient[i % len(soft_gradient)] for i in range(len(top_heroes))]
        bars = ax4.barh(top_heroes['英雄'], top_heroes['总胜率'], color=colors, edgecolor='white', linewidth=0.5)
        ax4.set_xlabel('胜率')
        ax4.set_title('英雄胜率TOP10（出场≥2次）')
        ax4.set_xlim(0, 1)
        ax4.invert_yaxis()
        
        # 添加数值标签
        for i, (bar, row) in enumerate(zip(bars, top_heroes.itertuples())):
            ax4.text(bar.get_width() + 0.01, bar.get_y() + bar.get_height()/2,
                    f'{row.总胜率:.2%}', ha='left', va='center', fontsize=9, color=CHART_COLORS['text'])
    
    plt.tight_layout()
    plt.savefig(f'{charts_dir}/英雄胜率TOP10.png', bbox_inches='tight', facecolor='white')
    plt.close()
    
    # 5. 最佳队友组合胜率TOP8图表
    fig5, ax5 = plt.subplots(figsize=(10, 6))
    best_teams = best_teammate_leaderboard.head(8).copy()
    if len(best_teams) > 0:
        best_teams = best_teams.sort_values(by='胜率',ascending=False)
        colors = [CHART_COLORS['success'] if i < 3 else CHART_COLORS['secondary'] for i in range(len(best_teams))]
        bars = ax5.barh(best_teams['队友组合'], best_teams['胜率'], color=colors, edgecolor='white', linewidth=0.5)
        ax5.set_xlabel('胜率')
        ax5.set_title('最佳队友组合TOP8')
        ax5.set_xlim(0, 1)
        ax5.invert_yaxis()
        
        # 添加数值标签
        for i, (bar, row) in enumerate(zip(bars, best_teams.itertuples())):
            ax5.text(bar.get_width() + 0.01, bar.get_y() + bar.get_height()/2,
                    f'{row.胜率:.2%} ({row.一起出场}场)', ha='left', va='center', fontsize=9, color=CHART_COLORS['text'])
    
    plt.tight_layout()
    plt.savefig(f'{charts_dir}/最佳队友组合TOP8.png', bbox_inches='tight', facecolor='white')
    plt.close()
    
    # 6. 每日比赛场次图表
    fig6, ax6 = plt.subplots(figsize=(10, 5))
    daily_data = daily_stats.copy()
    if len(daily_data) > 0:
        dates = daily_data['日期'].tolist()
        games = daily_data['比赛场次'].tolist()
        
        bars = ax6.bar(dates, games, color=CHART_COLORS['secondary'], edgecolor='white', linewidth=0.5)
        ax6.set_xlabel('日期')
        ax6.set_ylabel('比赛场次')
        ax6.set_title('每日比赛场次统计')
        ax6.tick_params(axis='x', rotation=45)
        
        # 添加数值标签
        for bar in bars:
            height = bar.get_height()
            ax6.text(bar.get_x() + bar.get_width()/2, height + 0.1,
                    f'{int(height)}', ha='center', va='bottom', fontsize=9, color=CHART_COLORS['text'])
    
    plt.tight_layout()
    plt.savefig(f'{charts_dir}/每日比赛场次.png', bbox_inches='tight', facecolor='white')
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
    
    # 获取当前时间（只显示日期）
    current_time = datetime.now().strftime("%Y-%m-%d")
    
    # 统计摘要
    total_games = len(df)
    total_players = len(player_stats)
    total_heroes = len(hero_stats)
    total_days = len(df['比赛时间'].unique())
    blue_wins = len(df[df['胜方'] == '蓝'])
    red_wins = len(df[df['胜方'] == '红'])
    blue_win_rate = blue_wins / total_games * 100 if total_games > 0 else 0
    red_win_rate = red_wins / total_games * 100 if total_games > 0 else 0
    
    # 获取关键数据
    top_players = player_leaderboard.head(10)
    top_heroes = hero_leaderboard.head(10)
    top_mvp = mvp_leaderboard.head(10)
    top_hero_pool = hero_pool_leaderboard.head(10)
    top_best_teammate = best_teammate_leaderboard.head(8)
    top_best_hero_combo = best_hero_combo_leaderboard.head(8)
    top_activity = player_activity_leaderboard.head(8)
    top_streak = streak_leaderboard.head(8)
    # 生成2026赏金榜
    bounty_leaderboard_2026, bounty_pool_info, bounty_dates = calculate_bounty_leaderboard(df_2026)
    print("赏金榜 (2026年):")
    print(bounty_leaderboard_2026)
    print(f"\n奖金池信息: 初始={bounty_pool_info['initial']}元, 已发放={bounty_pool_info['distributed']}元, 剩余={bounty_pool_info['remaining']}元")
    
    # 创建HTML内容
    html_content = f"""
    <!DOCTYPE html>
    <html lang="zh-CN">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>🏆 "饼干杯"-BYG王者荣耀联赛数据统计报告</title>
        <style>
            :root {{
                --primary-color: #5B7BA3;
                --primary-light: #7BA3C9;
                --primary-dark: #4A6890;
                --accent-color: #E8A87C;
                --accent-light: #F4D9C6;
                --success-color: #7CB798;
                --success-light: #C4E1D4;
                --warning-color: #E8C87C;
                --danger-color: #C98B8B;
                --danger-light: #E8D5D5;
                --text-primary: #4A5568;
                --text-secondary: #718096;
                --text-light: #A0AEC0;
                --bg-primary: #F7F9FC;
                --bg-secondary: #EDF2F7;
                --bg-card: #FFFFFF;
                --border-color: #E2E8F0;
                --border-light: #EDF2F7;
            }}
            
            * {{
                margin: 0;
                padding: 0;
                box-sizing: border-box;
            }}
            
            body {{
                font-family: 'Microsoft YaHei', Arial, sans-serif;
                font-size: 11px;
                line-height: 1.4;
                color: var(--text-primary);
                background-color: var(--bg-primary);
                padding: 8px;
            }}
            
            .container {{
                max-width: 100%;
                margin: 0 auto;
                background: var(--bg-card);
                border-radius: 8px;
                box-shadow: 0 2px 8px rgba(74, 85, 104, 0.1);
                overflow: hidden;
            }}
            
            /* 头部样式 - 柔和蓝灰渐变 */
            .header {{
                background: linear-gradient(135deg, var(--primary-color) 0%, var(--primary-dark) 100%);
                color: white;
                padding: 16px 20px;
                text-align: center;
                border-bottom: 3px solid var(--accent-color);
            }}
            
            .header h1 {{
                font-size: 18px;
                margin-bottom: 6px;
                font-weight: bold;
                text-shadow: 1px 1px 2px rgba(0,0,0,0.15);
            }}
            
            .header p {{
                font-size: 10px;
                opacity: 0.9;
            }}
            
            /* 统计卡片 */
            .stats-grid {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(120px, 1fr));
                gap: 8px;
                padding: 12px;
                background: var(--bg-secondary);
                border-bottom: 1px solid var(--border-color);
            }}
            
            .stat-card {{
                background: var(--bg-card);
                border-radius: 6px;
                padding: 10px;
                text-align: center;
                box-shadow: 0 1px 3px rgba(74, 85, 104, 0.08);
                border-top: 3px solid var(--primary-color);
                transition: transform 0.2s, box-shadow 0.2s;
            }}
            
            .stat-card:hover {{
                transform: translateY(-2px);
                box-shadow: 0 4px 12px rgba(74, 85, 104, 0.12);
            }}
            
            .stat-card .value {{
                font-size: 16px;
                font-weight: bold;
                color: var(--primary-dark);
                margin: 4px 0;
            }}
            
            .stat-card .label {{
                font-size: 9px;
                color: var(--text-secondary);
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }}
            
            /* 主要内容区 */
            .content {{
                padding: 12px;
            }}
            
            .section {{
                background: var(--bg-card);
                border-radius: 8px;
                padding: 14px;
                margin-bottom: 14px;
                box-shadow: 0 1px 4px rgba(74, 85, 104, 0.08);
                border: 1px solid var(--border-color);
                page-break-inside: avoid;
                transition: all 0.3s ease;
            }}
            
            .section:hover {{
                box-shadow: 0 4px 12px rgba(74, 85, 104, 0.12);
            }}
            
            .section-title {{
                font-size: 14px;
                font-weight: bold;
                color: var(--primary-color);
                margin-bottom: 10px;
                padding-bottom: 6px;
                border-bottom: 2px solid var(--accent-color);
                display: flex;
                justify-content: space-between;
                align-items: center;
            }}
            
            .section-title span {{
                font-size: 10px;
                font-weight: normal;
                color: var(--text-secondary);
            }}
            
            /* 表格样式 */
            .data-table {{
                width: 100%;
                border-collapse: collapse;
                font-size: 10px;
                margin-bottom: 10px;
                border: 1px solid var(--border-color);
                border-radius: 6px;
                overflow: hidden;
            }}
            
            .data-table th {{
                background: linear-gradient(to bottom, var(--bg-secondary), var(--border-light));
                color: var(--text-primary);
                text-align: left;
                padding: 8px 10px;
                font-weight: 600;
                border-bottom: 2px solid var(--border-color);
                font-size: 9px;
                text-transform: uppercase;
                letter-spacing: 0.3px;
                border-right: 1px solid var(--border-color);
            }}
            
            .data-table td {{
                padding: 6px 10px;
                border-bottom: 1px solid var(--border-light);
                border-right: 1px solid var(--border-light);
            }}
            
            .data-table tr:nth-child(even) {{
                background-color: var(--bg-primary);
            }}
            
            .data-table tr:hover {{
                background-color: rgba(91, 123, 163, 0.08);
            }}
            
            .data-table .rank-1 {{ 
                background-color: var(--accent-light) !important; 
                color: var(--primary-dark); 
                font-weight: bold; 
            }}
            .data-table .rank-2 {{ 
                background-color: var(--bg-secondary) !important;
                color: var(--accent-color); 
                font-weight: bold; 
            }}
            .data-table .rank-3 {{ 
                background-color: var(--bg-secondary) !important;
                color: var(--warning-color); 
                font-weight: bold; 
            }}
            
            /* 图表容器 */
            .charts-container {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
                gap: 10px;
                margin: 12px 0;
            }}
            
            .chart-box {{
                background: var(--bg-card);
                border-radius: 8px;
                padding: 12px;
                box-shadow: 0 1px 4px rgba(74, 85, 104, 0.08);
                border: 1px solid var(--border-color);
                text-align: center;
                transition: all 0.3s ease;
            }}
            
            .chart-box:hover {{
                box-shadow: 0 4px 12px rgba(74, 85, 104, 0.12);
            }}
            
            .chart-box img {{
                max-width: 100%;
                height: auto;
                border-radius: 6px;
            }}
            
            .chart-title {{
                font-size: 12px;
                font-weight: bold;
                color: var(--text-primary);
                margin-bottom: 8px;
                padding-bottom: 6px;
                border-bottom: 2px solid var(--accent-color);
            }}
            
            /* 排行榜容器 */
            .leaderboard-container {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
                gap: 10px;
                margin-top: 12px;
            }}
            
            .leaderboard-box {{
                background: var(--bg-secondary);
                border-radius: 8px;
                padding: 10px;
                border: 1px solid var(--border-color);
                transition: all 0.3s ease;
            }}
            
            .leaderboard-box:hover {{
                box-shadow: 0 3px 8px rgba(74, 85, 104, 0.1);
            }}
            
            .leaderboard-title {{
                font-size: 12px;
                font-weight: bold;
                color: var(--primary-color);
                margin-bottom: 8px;
                text-align: center;
                padding-bottom: 5px;
                border-bottom: 2px solid var(--accent-color);
            }}
            
            /* 胜率卡片（用于英雄-玩家和玩家-英雄排行榜） */
            .winrate-cards {{
                display: grid;
                grid-template-columns: repeat(auto-fill, minmax(170px, 1fr));
                gap: 8px;
                margin-top: 12px;
            }}
            
            .winrate-card {{
                background: var(--bg-card);
                border-radius: 8px;
                padding: 10px;
                border: 1px solid var(--border-color);
                box-shadow: 0 1px 3px rgba(74, 85, 104, 0.06);
                transition: all 0.3s ease;
            }}
            
            .winrate-card:hover {{
                box-shadow: 0 4px 12px rgba(74, 85, 104, 0.12);
                transform: translateY(-2px);
            }}
            
            .winrate-card .title {{
                font-size: 11px;
                font-weight: bold;
                color: var(--text-primary);
                margin-bottom: 6px;
                padding-bottom: 4px;
                border-bottom: 1px solid var(--accent-color);
                display: flex;
                justify-content: space-between;
                align-items: center;
            }}
            
            .winrate-card .item {{
                display: flex;
                justify-content: space-between;
                font-size: 9px;
                padding: 4px 0;
                border-bottom: 1px dashed var(--border-color);
            }}
            
            .winrate-card .item:last-child {{
                border-bottom: none;
            }}
            
            .winrate-card .item .name {{
                color: var(--text-primary);
            }}
            
            .winrate-card .item .rate {{
                color: var(--success-color);
                font-weight: bold;
            }}
            
            /* 分路排行榜 */
            .position-grid {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
                gap: 10px;
                margin-top: 12px;
            }}
            
            .position-box {{
                background: var(--bg-secondary);
                border-radius: 8px;
                padding: 10px;
                border: 1px solid var(--border-color);
                transition: all 0.3s ease;
            }}
            
            .position-box:hover {{
                box-shadow: 0 3px 8px rgba(74, 85, 104, 0.1);
            }}
            
            .position-title {{
                font-size: 11px;
                font-weight: bold;
                color: var(--primary-color);
                margin-bottom: 6px;
                text-align: center;
                padding-bottom: 5px;
                border-bottom: 2px solid var(--accent-color);
            }}
            
            /* 玩家详细数据区域 */
            .player-detail-grid {{
                display: grid;
                grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
                gap: 12px;
                margin-top: 12px;
            }}
            
            .player-detail-card {{
                background: linear-gradient(135deg, var(--bg-secondary) 0%, var(--bg-tertiary) 100%);
                border-radius: 8px;
                padding: 12px;
                border-left: 4px solid var(--primary-color);
                box-shadow: 0 1px 4px rgba(74, 85, 104, 0.08);
                transition: all 0.3s ease;
            }}
            
            .player-detail-card:hover {{
                box-shadow: 0 4px 12px rgba(74, 85, 104, 0.12);
                transform: translateY(-3px);
            }}
            
            .player-detail-card .name {{
                font-size: 13px;
                font-weight: bold;
                color: var(--text-primary);
                margin-bottom: 8px;
                display: flex;
                justify-content: space-between;
                align-items: center;
            }}
            
            .player-detail-card .stats-row {{
                display: flex;
                justify-content: space-between;
                font-size: 10px;
                color: var(--text-secondary);
                margin-bottom: 5px;
                padding: 3px 0;
                border-bottom: 1px dashed var(--border-color);
            }}
            
            .player-detail-card .stats-row:last-child {{
                border-bottom: none;
            }}
            
            .player-detail-card .hero-list {{
                font-size: 9px;
                color: var(--text-secondary);
                margin-top: 8px;
                line-height: 1.4;
            }}
            
            /* 英雄详细数据区域 */
            .hero-detail-grid {{
                display: grid;
                grid-template-columns: repeat(auto-fill, minmax(180px, 1fr));
                gap: 10px;
                margin-top: 12px;
            }}
            
            .hero-detail-card {{
                background: var(--bg-card);
                border-radius: 8px;
                padding: 12px;
                text-align: center;
                border-top: 3px solid var(--primary-color);
                box-shadow: 0 1px 4px rgba(74, 85, 104, 0.08);
                transition: all 0.3s ease;
            }}
            
            .hero-detail-card:hover {{
                box-shadow: 0 4px 12px rgba(74, 85, 104, 0.12);
                transform: translateY(-3px);
            }}
            
            .hero-detail-card .name {{
                font-size: 12px;
                font-weight: bold;
                color: var(--text-primary);
                margin-bottom: 6px;
            }}
            
            .hero-detail-card .win-rate {{
                font-size: 11px;
                color: var(--success-color);
                font-weight: bold;
                margin-bottom: 5px;
            }}
            
            /* 基础数据表格 */
            .base-data-container {{
                max-height: 400px;
                overflow-y: auto;
                margin-top: 12px;
                border: 1px solid var(--border-color);
                border-radius: 8px;
            }}
            
            /* 底部样式 */
            .footer {{
                text-align: center;
                padding: 12px;
                color: var(--text-secondary);
                font-size: 9px;
                border-top: 1px solid var(--border-color);
                background: var(--bg-secondary);
                margin-top: 20px;
                border-radius: 0 0 8px 8px;
            }}
            
            .copyright {{
                color: var(--primary-color);
                font-weight: bold;
                margin-top: 5px;
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
                
                .winrate-cards {{
                    grid-template-columns: 1fr;
                }}
                
                .player-detail-grid {{
                    grid-template-columns: 1fr;
                }}
                
                .hero-detail-grid {{
                    grid-template-columns: repeat(2, 1fr);
                }}
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
            
            /* 标签样式 */
            .tag {{
                display: inline-block;
                padding: 2px 6px;
                font-size: 8px;
                border-radius: 4px;
                margin-left: 5px;
                font-weight: bold;
            }}
            
            .tag-blue {{ background: var(--primary-color); color: white; }}
            .tag-red {{ background: var(--danger-color); color: white; }}
            .tag-green {{ background: var(--success-color); color: white; }}
            .tag-gold {{ background: var(--accent-color); color: var(--text-primary); }}
            .tag-orange {{ background: #D4956A; color: white; }}
            
            /* Tab导航样式 */
            .tab-nav {{
                display: flex;
                flex-wrap: wrap;
                background: var(--bg-secondary);
                border-radius: 8px 8px 0 0;
                padding: 8px 8px 0 8px;
                gap: 4px;
                border-bottom: 2px solid var(--primary-color);
            }}
            
            .tab-btn {{
                padding: 10px 16px;
                border: none;
                background: var(--bg-card);
                color: var(--text-secondary);
                font-size: 11px;
                font-weight: 600;
                cursor: pointer;
                border-radius: 6px 6px 0 0;
                transition: all 0.3s ease;
                border: 1px solid var(--border-color);
                border-bottom: none;
                margin-bottom: -2px;
            }}
            
            .tab-btn:hover {{
                background: var(--primary-light);
                color: white;
            }}
            
            .tab-btn.active {{
                background: var(--primary-color);
                color: white;
                border-color: var(--primary-color);
            }}
            
            .tab-content {{
                display: none;
                padding: 15px;
                animation: fadeIn 0.3s ease;
            }}
            
            .tab-content.active {{
                display: block;
            }}
            
            @keyframes fadeIn {{
                from {{ opacity: 0; transform: translateY(10px); }}
                to {{ opacity: 1; transform: translateY(0); }}
            }}
            
            .tab-container {{
                background: var(--bg-card);
                border-radius: 0 0 8px 8px;
                border: 1px solid var(--border-color);
                border-top: none;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <!-- 头部 -->
            <div class="header">
                <h1>🏆 "饼干杯"-BYG王者荣耀联赛数据统计报告</h1>
                <p>数据统计日期: {current_time} | 共 {total_games} 场比赛 | {total_players} 名玩家 | {total_heroes} 个英雄</p>
            </div>
            
            <!-- 关键统计 -->
            <div class="section">
                <div class="section-title">📊 关键统计数据</div>
                <div class="stats-grid">
                    <div class="stat-card">
                        <div class="value">{total_games}</div>
                        <div class="label">总比赛场次</div>
                    </div>
                    <div class="stat-card">
                        <div class="value">{total_days}</div>
                        <div class="label">总比赛天数</div>
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
                        <div class="value">{blue_win_rate:.1f}%</div>
                        <div class="label">蓝方胜率</div>
                    </div>
                    <div class="stat-card">
                        <div class="value">{red_wins}</div>
                        <div class="label">红方胜场</div>
                    </div>
                    <div class="stat-card">
                        <div class="value">{red_win_rate:.1f}%</div>
                        <div class="label">红方胜率</div>
                    </div>
                </div>
            </div>
            
            <!-- Tab导航 -->
            <div class="tab-nav">
                <button class="tab-btn active" onclick="showTab('overview')">📊 总览</button>
                <button class="tab-btn" onclick="showTab('players')">👑 玩家排行</button>
                <button class="tab-btn" onclick="showTab('heroes')">⚔️ 英雄数据</button>
                <button class="tab-btn" onclick="showTab('combos')">🤝 组合&分路</button>
                <button class="tab-btn" onclick="showTab('details')">📋 详细数据</button>
                <button class="tab-btn" onclick="showTab('records')">📅 比赛记录</button>
            </div>
            
            <div class="tab-container">
                <!-- Tab1: 总览 -->
                <div id="overview" class="tab-content active">
                    <!-- 可视化图表 -->
                    <div class="section">
                        <div class="section-title">📈 数据可视化图表</div>
                        <div class="charts-container">
    """
    
    # 添加图表
    if charts_dir and os.path.exists(charts_dir):
        # 选择图表
        important_charts = [
            ('玩家胜率TOP10.png', '玩家胜率TOP10'),
            ('英雄池数量TOP10.png', '英雄池TOP10'),
            ('MVP次数分布.png', 'MVP排行榜'),
            ('英雄胜率TOP10.png', '英雄胜率TOP10'),
            ('最佳队友组合TOP8.png', '最佳队友组合TOP8'),
            ('每日比赛场次.png', '每日比赛场次统计')
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
                </div>
                
                <!-- Tab2: 玩家排行 -->
                <div id="players" class="tab-content">
            
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
                            <th width="60">最长连胜</th>
                            <th width="60">当前连胜</th>
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
                            <td>{row['最长连胜']}</td>
                            <td>{row['当前连胜']}</td>
                        </tr>
        """
    
    html_content += """
                    </tbody>
                </table>
            </div>
                </div>
                
                <!-- Tab3: 英雄数据 -->
                <div id="heroes" class="tab-content">
            
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
                            <th width="70">使用玩家数</th>
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
                            <td>{row['使用玩家数']}</td>
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
    
    # 选择出场次数最多的前16个英雄
    hero_games = [(hero, hero_stats[hero]['总场次']) for hero in hero_stats.keys()]
    hero_games.sort(key=lambda x: x[1], reverse=True)
    
    for hero, games in hero_games[:16]:
        if hero in hero_player_leaderboard and hero_player_leaderboard[hero]:
            html_content += f"""
                    <div class="winrate-card">
                        <div class="title">{hero} <span style="font-size:9px;color:#666;">({games}场)</span></div>
            """
            
            for i, player_data in enumerate(hero_player_leaderboard[hero][:4], 1):
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
    
    # 选择出场次数最多的前16个玩家
    player_games = [(player, player_stats[player]['总场次']) for player in player_stats.keys()]
    player_games.sort(key=lambda x: x[1], reverse=True)

    for player, games in player_games[:16]:
        if player in player_hero_leaderboard and player_hero_leaderboard[player]:
            html_content += f"""
                    <div class="winrate-card">
                        <div class="title">{player} <span style="font-size:9px;color:#666;">({games}场)</span></div>
            """
            
            for i, hero_data in enumerate(player_hero_leaderboard[player][:4], 1):
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
                </div>
                
                <!-- Tab4: 组合&分路 -->
                <div id="combos" class="tab-content">
            
            <!-- MVP排行榜和英雄池排行榜 -->
            <div class="section">
                <div class="section-title">⭐ 关键个人排行榜</div>
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
                                    <th width="50">MVP率</th>
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
                                    <td>{row['MVP率']}</td>
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
                                    <th width="60">平均场次</th>
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
                    
                    <!-- 连胜排行榜 -->
                    <div class="leaderboard-box">
                        <div class="leaderboard-title">连胜排行榜</div>
                        <table class="data-table">
                            <thead>
                                <tr>
                                    <th width="30">排名</th>
                                    <th>玩家</th>
                                    <th width="50">最长连胜</th>
                                    <th width="50">当前连胜</th>
                                </tr>
                            </thead>
                            <tbody>
    """
    
    for idx, row in top_streak.iterrows():
        rank_class = f"rank-{idx}" if idx <= 3 else ""
        html_content += f"""
                                <tr class="{rank_class}">
                                    <td>{idx}</td>
                                    <td>{row['玩家']}</td>
                                    <td>{row['最长连胜']}</td>
                                    <td>{row['当前连胜']}</td>
                                </tr>
        """
    
    html_content += """
                            </tbody>
                        </table>
                    </div>
                    
                    <!-- 活跃度排行榜 -->
                    <div class="leaderboard-box">
                        <div class="leaderboard-title">活跃度排行榜</div>
                        <table class="data-table">
                            <thead>
                                <tr>
                                    <th width="30">排名</th>
                                    <th>玩家</th>
                                    <th width="50">活跃天数</th>
                                    <th width="50">场均比赛</th>
                                </tr>
                            </thead>
                            <tbody>
    """
    
    for idx, row in top_activity.iterrows():
        rank_class = f"rank-{idx}" if idx <= 3 else ""
        html_content += f"""
                                <tr class="{rank_class}">
                                    <td>{idx}</td>
                                    <td>{row['玩家']}</td>
                                    <td>{row['活跃天数']}</td>
                                    <td>{row['场均比赛']}</td>
                                </tr>
        """
    
    html_content += """
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
            
            <!-- 最佳组合排行榜 -->
            <div class="section">
                <div class="section-title">🤝 最佳组合排行榜</div>
                <div class="leaderboard-container">
                    <!-- 最佳队友组合 -->
                    <div class="leaderboard-box">
                        <div class="leaderboard-title">最佳队友组合</div>
                        <table class="data-table">
                            <thead>
                                <tr>
                                    <th width="30">排名</th>
                                    <th>队友组合</th>
                                    <th width="40">场次</th>
                                    <th width="50">胜率</th>
                                </tr>
                            </thead>
                            <tbody>
    """
    
    for idx, row in top_best_teammate.iterrows():
        rank_class = f"rank-{idx}" if idx <= 3 else ""
        html_content += f"""
                                <tr class="{rank_class}">
                                    <td>{idx}</td>
                                    <td>{row['队友组合']}</td>
                                    <td>{row['一起出场']}</td>
                                    <td>{row['胜率百分比']}</td>
                                </tr>
        """
    
    html_content += """
                            </tbody>
                        </table>
                    </div>
                    
                    <!-- 最佳英雄组合 -->
                    <div class="leaderboard-box">
                        <div class="leaderboard-title">最佳英雄组合</div>
                        <table class="data-table">
                            <thead>
                                <tr>
                                    <th width="30">排名</th>
                                    <th>英雄组合</th>
                                    <th width="40">场次</th>
                                    <th width="50">胜率</th>
                                </tr>
                            </thead>
                            <tbody>
    """
    
    for idx, row in top_best_hero_combo.iterrows():
        rank_class = f"rank-{idx}" if idx <= 3 else ""
        html_content += f"""
                                <tr class="{rank_class}">
                                    <td>{idx}</td>
                                    <td>{row['英雄组合']}</td>
                                    <td>{row['一起出场']}</td>
                                    <td>{row['胜率百分比']}</td>
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
                        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 6px;">
                            <div>
                                <div style="font-size:10px;color:#5B7BA3;margin-bottom:4px;font-weight:bold;text-align:center;">玩家TOP5</div>
        """

        # 玩家TOP5
        for i, (_, row) in enumerate(pos_player_df.head(5).iterrows(), 1):
            html_content += f"""
                                <div style="font-size:10px;padding:3px 0;border-bottom:1px solid #E2E8F0;">
                                    <span>{i}. {row['玩家']}</span>
                                    <span style="float:right;color:#7CB798;font-weight:bold;">{row['胜率百分比']}</span>
                                </div>
            """
        
        html_content += """
                            </div>
                            <div>
                                <div style="font-size:10px;color:#5B7BA3;margin-bottom:4px;font-weight:bold;text-align:center;">英雄TOP5</div>
        """

        # 英雄TOP5
        for i, (_, row) in enumerate(pos_hero_df.head(5).iterrows(), 1):
            html_content += f"""
                                <div style="font-size:10px;padding:3px 0;border-bottom:1px solid #E2E8F0;">
                                    <span>{i}. {row['英雄']}</span>
                                    <span style="float:right;color:#7CB798;font-weight:bold;">{row['胜率百分比']}</span>
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
                </div>
                
                <!-- Tab5: 详细数据 -->
                <div id="details" class="tab-content">
            
            <!-- 玩家详细数据 -->
            <div class="section">
                <div class="section-title">👤 玩家详细数据</div>
                <div class="player-detail-grid">
    """
    
    # 添加所有玩家详细数据
    sorted_players = sorted(player_stats.items(), key=lambda x: x[1]['总场次'], reverse=True)
    
    for player, stats in sorted_players[:20]:  # 显示前20名玩家
        if stats['总场次'] > 0:
            win_rate = stats['总胜场'] / stats['总场次'] * 100
            
            # 统计各分路场次
            position_stats = []
            for position in positions:
                if stats[f'{position}场次'] > 0:
                    pos_win_rate = stats[f'{position}胜场'] / stats[f'{position}场次'] * 100 if stats[f'{position}场次'] > 0 else 0
                    position_stats.append(f"{position[:2]}: {stats[f'{position}场次']}场 ({pos_win_rate:.1f}%)")
            
            # 获取最擅长的英雄（胜率最高且至少出场2次）
            best_heroes = []
            for hero in stats['英雄场次']:
                if stats['英雄场次'][hero] >= 2 and stats['英雄胜场'][hero] > 0:
                    hero_win_rate = stats['英雄胜场'][hero] / stats['英雄场次'][hero] * 100
                    best_heroes.append((hero, hero_win_rate, stats['英雄场次'][hero]))
            
            best_heroes.sort(key=lambda x: x[1], reverse=True)
            top_heroes_str = ", ".join([f"{h[0]}({h[2]})" for h in best_heroes[:3]]) if best_heroes else "无"
            
            html_content += f"""
                    <div class="player-detail-card">
                        <div class="name">
                            {player}
                            <span style="font-size:11px;color:#5B7BA3;">{win_rate:.1f}%</span>
                        </div>
                        <div class="stats-row">
                            <span>总场次: {stats['总场次']}</span>
                            <span>胜场: {stats['总胜场']}</span>
                        </div>
                        <div class="stats-row">
                            <span>MVP: {stats['MVP次数']}</span>
                            <span>英雄池: {len(stats['英雄池'])}</span>
                        </div>
                        <div class="stats-row">
                            <span>最长连胜: {stats['最长连胜']}</span>
                            <span>当前连胜: {stats['连胜']}</span>
                        </div>
                        <div class="stats-row">
                            <span>分路分布:</span>
                        </div>
                        <div class="stats-row" style="font-size:9px;">
                            {', '.join(position_stats)}
                        </div>
                        <div class="stats-row">
                            <span>擅长英雄:</span>
                        </div>
                        <div class="hero-list">
                            {top_heroes_str}
                        </div>
                    </div>
            """
    
    html_content += """
                </div>
            </div>
            
            <!-- 英雄详细数据 -->
            <div class="section">
                <div class="section-title">⚔️ 英雄详细数据</div>
                <div class="hero-detail-grid">
    """
    
    # 添加所有英雄详细数据
    sorted_heroes = sorted(hero_stats.items(), key=lambda x: x[1]['总场次'], reverse=True)
    
    for hero, stats in sorted_heroes[:30]:  # 显示前30个英雄
        if stats['总场次'] > 0:
            win_rate = stats['总胜场'] / stats['总场次'] * 100
            
            # 统计主要分路
            main_position = ""
            max_games = 0
            for position in positions:
                if stats[f'{position}场次'] > max_games:
                    max_games = stats[f'{position}场次']
                    main_position = position
            
            # 获取最佳使用者（胜率最高且至少出场2次）
            best_players = []
            for player in stats['玩家场次']:
                if stats['玩家场次'][player] >= 2 and stats['玩家胜场'][player] > 0:
                    player_win_rate = stats['玩家胜场'][player] / stats['玩家场次'][player] * 100
                    best_players.append((player, player_win_rate, stats['玩家场次'][player]))
            
            best_players.sort(key=lambda x: x[1], reverse=True)
            top_players_str = ", ".join([f"{p[0]}({p[2]})" for p in best_players[:2]]) if best_players else "无"
            
            html_content += f"""
                    <div class="hero-detail-card">
                        <div class="name">{hero}</div>
                        <div class="win-rate">{win_rate:.1f}%</div>
                        <div style="font-size:9px;color:#666;margin-bottom:3px;">出场: {stats['总场次']}次</div>
                        <div style="font-size:8px;color:#999;margin-bottom:3px;">主要分路: {main_position}</div>
                        <div style="font-size:8px;color:#777;">使用玩家: {len(stats['玩家场次'])}人</div>
                        <div style="font-size:8px;color:#555;margin-top:4px;">最佳使用者: {top_players_str}</div>
                    </div>
            """
    
    html_content += """
                </div>
            </div>
                </div>
                
                <!-- Tab6: 比赛记录 -->
                <div id="records" class="tab-content">
            
            <!-- 2026年赏金榜 -->
            <div class="section" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 16px; padding: 20px; margin-bottom: 20px;">
                <div style="display: flex; justify-content: space-around; flex-wrap: wrap; gap: 15px; margin-bottom: 10px;">
                    <div style="text-align: center; color: white;">
                        <div style="font-size: 11px; opacity: 0.9; margin-bottom: 5px;">💰 奖金池初始值</div>
                        <div style="font-size: 24px; font-weight: bold; text-shadow: 0 2px 4px rgba(0,0,0,0.2);">""" + f"{bounty_pool_info['initial']}" + """ 元</div>
                    </div>
                    <div style="text-align: center; color: white;">
                        <div style="font-size: 11px; opacity: 0.9; margin-bottom: 5px;">📤 已发放</div>
                        <div style="font-size: 24px; font-weight: bold; text-shadow: 0 2px 4px rgba(0,0,0,0.2);">""" + f"{bounty_pool_info['distributed']}" + """ 元</div>
                    </div>
                    <div style="text-align: center; color: white;">
                        <div style="font-size: 11px; opacity: 0.9; margin-bottom: 5px;">💎 剩余</div>
                        <div style="font-size: 24px; font-weight: bold; text-shadow: 0 2px 4px rgba(0,0,0,0.2); """ + ("color: #ffcccb;" if bounty_pool_info['remaining'] < 0 else ("color: #ffeaa7;" if bounty_pool_info['remaining'] < 100 else "")) + """">""" + f"{bounty_pool_info['remaining']}" + """ 元</div>
                    </div>
                </div>
            </div>
            
            <div class="section">
                <div class="section-title">💰 2026赏金榜 <span style="font-size:10px;font-weight:normal;color:#666;">胜方全员+2元，胜方MVP+1元，败方MVP+1元</span></div>
                <table class="data-table">
                    <thead>
                        <tr>
                            <th width="30">排名</th>
                            <th>玩家</th>
                            <th width="50">总赏金</th>
    """
    
    # 添加每日日期列头
    for date in bounty_dates:
        short_date = date[5:]  # 从 "2026-01-02" 变成 "01-02"
        html_content += f"""
                            <th width="45">{short_date}</th>
        """
    
    html_content += """
                        </tr>
                    </thead>
                    <tbody>
    """
    
    # 添加赏金榜数据行
    for idx, row in bounty_leaderboard_2026.iterrows():
        rank_class = f"rank-{idx}" if idx <= 3 else ""
        html_content += f"""
                        <tr class="{rank_class}">
                            <td>{idx}</td>
                            <td>{row['玩家']}</td>
                            <td><strong>{row['总赏金']:.0f}</strong></td>
        """
        for date in bounty_dates:
            daily_val = row.get(date, 0)
            cell_style = "color:#999;" if daily_val == 0 else ""
            html_content += f"""
                            <td style="{cell_style}">{daily_val:.0f}</td>
            """
        html_content += """
                        </tr>
        """
    
    html_content += """
                    </tbody>
                </table>
            </div>
            
            <!-- 每日比赛统计 -->
            <div class="section">
                <div class="section-title">📅 每日比赛统计</div>
                <table class="data-table">
                    <thead>
                        <tr>
                            <th width="40">序号</th>
                            <th width="90">日期</th>
                            <th width="60">比赛场次</th>
                            <th width="60">蓝方胜场</th>
                            <th width="60">红方胜场</th>
                            <th width="70">蓝方胜率</th>
                            <th width="70">红方胜率</th>
                        </tr>
                    </thead>
                    <tbody>
    """
    
    for idx, row in daily_stats.iterrows():
        html_content += f"""
                        <tr>
                            <td>{idx}</td>
                            <td>{row['日期']}</td>
                            <td>{row['比赛场次']}</td>
                            <td>{row['蓝方胜场']}</td>
                            <td>{row['红方胜场']}</td>
                            <td>{row['蓝方胜率']}</td>
                            <td>{row['红方胜率']}</td>
                        </tr>
        """
    
    html_content += """
                    </tbody>
                </table>
            </div>
            
            <!-- 基础数据 - 完整比赛记录 -->
            <div class="section">
                <div class="section-title">📋 完整比赛记录</div>
                <div class="base-data-container">
                    <table class="data-table">
                        <thead>
                            <tr>
                                <th width="50">比赛ID</th>
                                <th width="80">日期</th>
                                <th width="40">胜方</th>
                                <th>蓝方边路</th>
                                <th>蓝方打野</th>
                                <th>蓝方中路</th>
                                <th>蓝方发育路</th>
                                <th>蓝方游走</th>
                                <th width="50">蓝方MVP</th>
                                <th>红方边路</th>
                                <th>红方打野</th>
                                <th>红方中路</th>
                                <th>红方发育路</th>
                                <th>红方游走</th>
                                <th width="50">红方MVP</th>
                            </tr>
                        </thead>
                        <tbody>
    """
    
    for idx, row in df.iterrows():
        # 格式化日期，只显示日期部分
        match_date = row['比赛时间']
        if hasattr(match_date, 'strftime'):
            match_date_str = match_date.strftime('%Y-%m-%d')
        else:
            match_date_str = str(match_date).split(' ')[0] if ' ' in str(match_date) else str(match_date)
        
        html_content += f"""
                            <tr>
                                <td>{row['比赛ID']}</td>
                                <td>{match_date_str}</td>
                                <td><span class="tag tag-blue">{row['胜方']}</span></td>
                                <td>{row['蓝方边']}</td>
                                <td>{row['蓝方野']}</td>
                                <td>{row['蓝方中']}</td>
                                <td>{row['蓝方射']}</td>
                                <td>{row['蓝方辅']}</td>
                                <td>{row['蓝方MVP']}</td>
                                <td>{row['红方边']}</td>
                                <td>{row['红方野']}</td>
                                <td>{row['红方中']}</td>
                                <td>{row['红方射']}</td>
                                <td>{row['红方辅']}</td>
                                <td>{row['红方MVP']}</td>
                            </tr>
        """
    
    html_content += f"""
                        </tbody>
                    </table>
                </div>
            </div>
                </div>
            </div>
            
            <!-- 底部信息 -->
            <div class="footer">
                <p>报告生成日期: {current_time} | 数据来源: {total_games}场BYG内战赛 | 统计玩家: {total_players}人 | 统计英雄: {total_heroes}个 | 比赛天数: {total_days}天</p>
                <p class="copyright">Copyright: Yuanhang Zhang -- v2.0</p>
                <p style="font-size:8px;color:#888;margin-top:5px;">🏆 "饼干杯"-BYG王者荣耀联赛数据统计报告 - 专业数据分析系统</p>
            </div>
        </div>
        
        <script>
            function showTab(tabId) {{
                // 隐藏所有tab内容
                document.querySelectorAll('.tab-content').forEach(tab => {{
                    tab.classList.remove('active');
                }});
                
                // 移除所有按钮的active状态
                document.querySelectorAll('.tab-btn').forEach(btn => {{
                    btn.classList.remove('active');
                }});
                
                // 显示选中的tab
                document.getElementById(tabId).classList.add('active');
                
                // 给对应按钮添加active状态
                event.target.classList.add('active');
            }}
        </script>
    </body>
    </html>
    """
    
    # 保存HTML文件
    with open('饼干杯-BYG王者荣耀联赛数据统计报告.html', 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print("完整HTML报告已生成: 饼干杯-BYG王者荣耀联赛数据统计报告.html")

# 生成Excel数据文件
def generate_excel_data():
    """生成包含所有数据的Excel文件"""
    print("正在生成Excel数据文件...")
    
    with pd.ExcelWriter('饼干杯-BYG王者荣耀联赛数据.xlsx', engine='openpyxl') as writer:
        # 原始数据
        df.to_excel(writer, sheet_name='原始比赛数据', index=False)
        
        # 核心排行榜
        player_leaderboard.to_excel(writer, sheet_name='玩家综合排行榜')
        hero_leaderboard.to_excel(writer, sheet_name='英雄胜率排行榜')
        mvp_leaderboard.to_excel(writer, sheet_name='MVP排行榜')
        hero_pool_leaderboard.to_excel(writer, sheet_name='英雄池排行榜')
        best_teammate_leaderboard.to_excel(writer, sheet_name='最佳队友组合榜')
        best_hero_combo_leaderboard.to_excel(writer, sheet_name='最佳英雄组合榜')
        daily_stats.to_excel(writer, sheet_name='每日比赛统计')
        player_activity_leaderboard.to_excel(writer, sheet_name='玩家活跃度榜')
        streak_leaderboard.to_excel(writer, sheet_name='连胜排行榜')
        
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
                # 计算各分路胜率
                position_winrates = {}
                for position in positions:
                    if stats[f'{position}场次'] > 0:
                        winrate = stats[f'{position}胜场'] / stats[f'{position}场次'] * 100
                        position_winrates[f'{position}胜率'] = f"{winrate:.1f}%"
                    else:
                        position_winrates[f'{position}胜率'] = "0%"
                
                player_detail_data.append({
                    '玩家': player,
                    '总场次': stats['总场次'],
                    '总胜场': stats['总胜场'],
                    '总胜率': f"{stats['总胜场']/stats['总场次']*100:.2f}%" if stats['总场次'] > 0 else "0%",
                    'MVP次数': stats['MVP次数'],
                    '英雄池数量': len(stats['英雄池']),
                    '最长连胜': stats['最长连胜'],
                    '当前连胜': stats['连胜'],
                    '活跃天数': len(stats['每日比赛']),
                    '边路场次': stats['边路场次'],
                    '打野场次': stats['打野场次'],
                    '中路场次': stats['中路场次'],
                    '发育路场次': stats['发育路场次'],
                    '游走场次': stats['游走场次'],
                    **position_winrates,
                    '英雄池列表': ', '.join(sorted(stats['英雄池']))
                })
        
        player_detail_df = pd.DataFrame(player_detail_data)
        player_detail_df.to_excel(writer, sheet_name='玩家详细数据', index=False)
        
        # 英雄详细数据
        hero_detail_data = []
        for hero, stats in hero_stats.items():
            if stats['总场次'] > 0:
                # 计算各分路胜率
                position_winrates = {}
                for position in positions:
                    if stats[f'{position}场次'] > 0:
                        winrate = stats[f'{position}胜场'] / stats[f'{position}场次'] * 100
                        position_winrates[f'{position}胜率'] = f"{winrate:.1f}%"
                    else:
                        position_winrates[f'{position}胜率'] = "0%"
                
                hero_detail_data.append({
                    '英雄': hero,
                    '总场次': stats['总场次'],
                    '总胜场': stats['总胜场'],
                    '总胜率': f"{stats['总胜场']/stats['总场次']*100:.2f}%" if stats['总场次'] > 0 else "0%",
                    '使用玩家数': len(stats['玩家场次']),
                    '边路场次': stats['边路场次'],
                    '打野场次': stats['打野场次'],
                    '中路场次': stats['中路场次'],
                    '发育路场次': stats['发育路场次'],
                    '游走场次': stats['游走场次'],
                    **position_winrates,
                    '最佳玩家': max(stats['玩家胜场'].items(), key=lambda x: x[1])[0] if stats['玩家胜场'] else "无"
                })
        
        hero_detail_df = pd.DataFrame(hero_detail_data)
        hero_detail_df.to_excel(writer, sheet_name='英雄详细数据', index=False)
        
        # 同一个英雄，玩家胜率榜
        hero_player_data = []
        for hero, player_list in hero_player_leaderboard.items():
            for player_data in player_list[:10]:  # 每个英雄取前10名
                hero_player_data.append({
                    '英雄': hero,
                    '玩家': player_data['玩家'],
                    '场次': player_data['场次'],
                    '胜场': player_data['胜场'],
                    '胜率': player_data['胜率百分比']
                })
        
        if hero_player_data:
            hero_player_df = pd.DataFrame(hero_player_data)
            hero_player_df.to_excel(writer, sheet_name='英雄玩家胜率榜', index=False)
        
        # 同一个玩家，英雄胜率榜
        player_hero_data = []
        for player, hero_list in player_hero_leaderboard.items():
            for hero_data in hero_list[:10]:  # 每个玩家取前10名
                player_hero_data.append({
                    '玩家': player,
                    '英雄': hero_data['英雄'],
                    '场次': hero_data['场次'],
                    '胜场': hero_data['胜场'],
                    '胜率': hero_data['胜率百分比']
                })
        
        if player_hero_data:
            player_hero_df = pd.DataFrame(player_hero_data)
            player_hero_df.to_excel(writer, sheet_name='玩家英雄胜率榜', index=False)
    
    print("Excel数据文件已生成: 饼干杯-BYG王者荣耀联赛数据.xlsx")

# 生成PDF报告（可选）
def generate_pdf_report():
    """生成PDF报告（需要weasyprint）"""
    try:
        from weasyprint import HTML
        
        print("正在生成PDF报告...")
        
        # 读取HTML内容
        with open('饼干杯-BYG王者荣耀联赛数据统计报告.html', 'r', encoding='utf-8') as f:
            html_content = f.read()
        
        # 生成PDF
        HTML(string=html_content).write_pdf('饼干杯-BYG王者荣耀联赛数据统计报告.pdf')
        
        print("PDF报告已生成: 饼干杯-BYG王者荣耀联赛数据统计报告.pdf")
        
    except ImportError:
        print("警告: weasyprint库未安装，无法生成PDF报告")
        print("请使用以下命令安装: pip install weasyprint")
    except Exception as e:
        print(f"PDF生成错误: {e}")

# 主程序
def main():
    print("\n" + "="*80)
    print("🏆 饼干杯-BYG王者荣耀联赛数据统计系统")
    print("="*80)
    
    # 生成所有报告
    generate_html_report()
    generate_excel_data()
    generate_pdf_report()
    
    print("\n" + "="*80)
    print("报告生成完成！")
    print("="*80)
    print("\n生成的文件:")
    print("1. 饼干杯-BYG王者荣耀联赛数据统计报告.html - 完整HTML报告")
    print("2. 饼干杯-BYG王者荣耀联赛数据.xlsx - 所有数据的Excel文件")
    print("3. 饼干杯-BYG王者荣耀联赛数据统计报告.pdf - PDF格式报告（如果已安装weasyprint）")
    print("4. charts/ - 包含所有可视化图表的目录")
    print("\n打开 饼干杯-BYG王者荣耀联赛数据统计报告.html 查看完整报告")

# 运行主程序
if __name__ == "__main__":
    main()