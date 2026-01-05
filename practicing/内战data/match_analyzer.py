"""
王者荣耀内战数据分析模块
将计算逻辑与展示逻辑解耦

核心函数: calculate_all_stats(df) - 接收比赛DataFrame，返回所有统计结果
"""

import pandas as pd
import numpy as np
from collections import defaultdict
from itertools import combinations
from datetime import datetime
import warnings

warnings.filterwarnings('ignore')

# ========== 常量定义 ==========
POSITIONS = ['边路', '打野', '中路', '发育路', '游走']

POSITION_MAP = {
    '蓝方边': '边路', '蓝方野': '打野', '蓝方中': '中路', '蓝方射': '发育路', '蓝方辅': '游走',
    '红方边': '边路', '红方野': '打野', '红方中': '中路', '红方射': '发育路', '红方辅': '游走'
}

MVP_POSITION_MAP = {'边': '边路', '野': '打野', '中': '中路', '射': '发育路', '辅': '游走'}
MVP_INDEX_MAP = {'边': 0, '野': 1, '中': 2, '射': 3, '辅': 4}

COLUMNS = ['比赛ID', '比赛时间', '胜方', '蓝方边', '蓝方野', '蓝方中', '蓝方射', '蓝方辅', '蓝方MVP', 
           '红方边', '红方野', '红方中', '红方射', '红方辅', '红方MVP']


# ========== 核心计算函数 ==========
def calculate_all_stats(df_input: pd.DataFrame) -> dict:
    """
    核心计算函数：基于比赛DataFrame计算所有统计数据
    
    参数:
        df_input: 包含比赛数据的DataFrame，需要包含以下列:
            - 比赛ID, 比赛时间, 胜方
            - 蓝方边, 蓝方野, 蓝方中, 蓝方射, 蓝方辅, 蓝方MVP
            - 红方边, 红方野, 红方中, 红方射, 红方辅, 红方MVP
    
    返回:
        dict: 包含所有统计结果的字典，包括:
            - basic_stats: 基础统计（比赛数、玩家数、英雄数等）
            - player_stats: 玩家详细统计字典
            - hero_stats: 英雄详细统计字典
            - player_leaderboard: 玩家综合排行榜DataFrame
            - hero_leaderboard: 英雄胜率排行榜DataFrame
            - mvp_leaderboard: MVP排行榜DataFrame
            - hero_pool_leaderboard: 英雄池排行榜DataFrame
            - streak_leaderboard: 连胜排行榜DataFrame
            - activity_leaderboard: 活跃度排行榜DataFrame
            - teammate_leaderboard: 最佳队友组合排行榜DataFrame
            - hero_combo_leaderboard: 最佳英雄组合排行榜DataFrame
            - daily_stats: 每日统计DataFrame
            - position_leaderboards: 各分路排行榜字典
            - hero_player_leaderboard: 英雄-玩家胜率榜字典
            - player_hero_leaderboard: 玩家-英雄胜率榜字典
            - bounty_leaderboard: 赏金榜dict (含leaderboard, pool_info, dates等)
            - df: 原始数据DataFrame
    """
    if df_input is None or len(df_input) == 0:
        return None
    
    df = df_input.copy()
    
    # 确保比赛时间是字符串格式
    if df['比赛时间'].dtype != 'object':
        df['比赛时间'] = df['比赛时间'].apply(lambda x: x.strftime("%Y-%m-%d") if hasattr(x, 'strftime') else str(x))
    
    # ========== 初始化统计字典 ==========
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
        '每日比赛': defaultdict(int),
        '连胜': 0,
        '最长连胜': 0,
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
    })
    
    player_pair_stats = defaultdict(lambda: {'一起出场': 0, '一起胜利': 0})
    hero_pair_stats = defaultdict(lambda: {'一起出场': 0, '一起胜利': 0})
    
    # ========== 遍历每场比赛进行统计 ==========
    for idx, row in df.iterrows():
        winner = row['胜方']
        match_date = row['比赛时间']
        
        # 处理蓝方
        blue_players, blue_heroes = [], []
        for pos_col in ['蓝方边', '蓝方野', '蓝方中', '蓝方射', '蓝方辅']:
            cell = row[pos_col]
            if isinstance(cell, str) and '-' in cell:
                player, hero = cell.split('-', 1)
                position = POSITION_MAP[pos_col]
                is_win = (winner == '蓝')
                
                _update_player_stats(player_stats[player], position, hero, match_date, is_win)
                _update_hero_stats(hero_stats[hero], position, player, is_win)
                
                blue_players.append(player)
                blue_heroes.append(hero)
        
        # 处理红方
        red_players, red_heroes = [], []
        for pos_col in ['红方边', '红方野', '红方中', '红方射', '红方辅']:
            cell = row[pos_col]
            if isinstance(cell, str) and '-' in cell:
                player, hero = cell.split('-', 1)
                position = POSITION_MAP[pos_col]
                is_win = (winner == '红')
                
                _update_player_stats(player_stats[player], position, hero, match_date, is_win)
                _update_hero_stats(hero_stats[hero], position, player, is_win)
                
                red_players.append(player)
                red_heroes.append(hero)
        
        # MVP统计
        _update_mvp_stats(player_stats, row['蓝方MVP'], blue_players)
        _update_mvp_stats(player_stats, row['红方MVP'], red_players)
        
        # 队友组合统计
        _update_pair_stats(player_pair_stats, blue_players, winner == '蓝')
        _update_pair_stats(player_pair_stats, red_players, winner == '红')
        
        # 英雄组合统计
        _update_pair_stats(hero_pair_stats, blue_heroes, winner == '蓝')
        _update_pair_stats(hero_pair_stats, red_heroes, winner == '红')
    
    # ========== 计算连胜数据 ==========
    _calculate_streaks(df, player_stats)
    
    # ========== 生成各种排行榜 ==========
    result = {
        'df': df,
        'player_stats': dict(player_stats),
        'hero_stats': dict(hero_stats),
        'basic_stats': _create_basic_stats(df, player_stats, hero_stats),
        'player_leaderboard': _create_player_leaderboard(player_stats),
        'hero_leaderboard': _create_hero_leaderboard(hero_stats),
        'mvp_leaderboard': _create_mvp_leaderboard(player_stats),
        'hero_pool_leaderboard': _create_hero_pool_leaderboard(player_stats),
        'streak_leaderboard': _create_streak_leaderboard(player_stats),
        'activity_leaderboard': _create_activity_leaderboard(player_stats),
        'teammate_leaderboard': _create_teammate_leaderboard(player_pair_stats),
        'hero_combo_leaderboard': _create_hero_combo_leaderboard(hero_pair_stats),
        'daily_stats': _create_daily_stats(df),
        'position_leaderboards': _create_position_leaderboards(player_stats, hero_stats),
        'hero_player_leaderboard': _create_hero_player_leaderboard(hero_stats),
        'player_hero_leaderboard': _create_player_hero_leaderboard(player_stats),
        'bounty_leaderboard': _create_bounty_leaderboard(df),
    }
    
    return result


# ========== 辅助更新函数 ==========
def _update_player_stats(stats: dict, position: str, hero: str, match_date: str, is_win: bool):
    """更新单个玩家的统计数据"""
    stats['总场次'] += 1
    stats[f'{position}场次'] += 1
    stats['英雄池'].add(hero)
    stats[f'{position}英雄池'].add(hero)
    stats['英雄场次'][hero] += 1
    stats['每日比赛'][match_date] += 1
    
    if is_win:
        stats['总胜场'] += 1
        stats[f'{position}胜场'] += 1
        stats['英雄胜场'][hero] += 1


def _update_hero_stats(stats: dict, position: str, player: str, is_win: bool):
    """更新单个英雄的统计数据"""
    stats['总场次'] += 1
    stats[f'{position}场次'] += 1
    stats['玩家场次'][player] += 1
    
    if is_win:
        stats['总胜场'] += 1
        stats[f'{position}胜场'] += 1
        stats['玩家胜场'][player] += 1


def _update_mvp_stats(player_stats: dict, mvp_pos: str, players: list):
    """更新MVP统计"""
    if mvp_pos in MVP_INDEX_MAP:
        mvp_idx = MVP_INDEX_MAP[mvp_pos]
        if mvp_idx < len(players):
            player_stats[players[mvp_idx]]['MVP次数'] += 1


def _update_pair_stats(pair_stats: dict, items: list, is_win: bool):
    """更新组合统计（队友或英雄）"""
    for item1, item2 in combinations(items, 2):
        key = tuple(sorted([item1, item2]))
        pair_stats[key]['一起出场'] += 1
        if is_win:
            pair_stats[key]['一起胜利'] += 1


def _calculate_streaks(df: pd.DataFrame, player_stats: dict):
    """计算玩家连胜数据"""
    df_sorted = df.sort_values('比赛时间')
    
    for idx, row in df_sorted.iterrows():
        winner = row['胜方']
        
        # 获取蓝方和红方玩家
        for side, win_condition in [('蓝', '蓝'), ('红', '红')]:
            prefix = f'{side}方'
            players = []
            for pos in ['边', '野', '中', '射', '辅']:
                cell = row[f'{prefix}{pos}']
                if isinstance(cell, str) and '-' in cell:
                    player, _ = cell.split('-', 1)
                    players.append(player)
            
            for player in players:
                if winner == win_condition:
                    player_stats[player]['连胜'] += 1
                else:
                    player_stats[player]['连胜'] = 0
                
                if player_stats[player]['连胜'] > player_stats[player]['最长连胜']:
                    player_stats[player]['最长连胜'] = player_stats[player]['连胜']


# ========== 排行榜生成函数 ==========
def _create_basic_stats(df: pd.DataFrame, player_stats: dict, hero_stats: dict) -> dict:
    """生成基础统计数据"""
    total_games = len(df)
    blue_wins = len(df[df['胜方'] == '蓝'])
    red_wins = len(df[df['胜方'] == '红'])
    
    return {
        'total_games': total_games,
        'total_days': len(df['比赛时间'].unique()),
        'total_players': len(player_stats),
        'total_heroes': len(hero_stats),
        'blue_wins': blue_wins,
        'red_wins': red_wins,
        'blue_win_rate': blue_wins / total_games * 100 if total_games > 0 else 0,
        'red_win_rate': red_wins / total_games * 100 if total_games > 0 else 0,
        'date_range': {
            'start': df['比赛时间'].min(),
            'end': df['比赛时间'].max()
        }
    }


def _create_player_leaderboard(player_stats: dict) -> pd.DataFrame:
    """生成玩家综合排行榜"""
    data = []
    for player, stats in player_stats.items():
        if stats['总场次'] > 0:
            win_rate = stats['总胜场'] / stats['总场次']
            data.append({
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
    
    df = pd.DataFrame(data)
    if len(df) > 0:
        df = df.sort_values(by=['总胜率', '总场次'], ascending=[False, False]).reset_index(drop=True)
        df.index = df.index + 1
    return df


def _create_hero_leaderboard(hero_stats: dict) -> pd.DataFrame:
    """生成英雄胜率排行榜"""
    data = []
    for hero, stats in hero_stats.items():
        if stats['总场次'] > 0:
            win_rate = stats['总胜场'] / stats['总场次']
            data.append({
                '英雄': hero,
                '总场次': stats['总场次'],
                '总胜场': stats['总胜场'],
                '总胜率': win_rate,
                '总胜率百分比': f"{win_rate * 100:.2f}%",
                '使用玩家数': len(stats['玩家场次'])
            })
    
    df = pd.DataFrame(data)
    if len(df) > 0:
        df = df.sort_values(by=['总胜率', '总场次'], ascending=[False, False]).reset_index(drop=True)
        df.index = df.index + 1
    return df


def _create_mvp_leaderboard(player_stats: dict) -> pd.DataFrame:
    """生成MVP排行榜"""
    data = []
    for player, stats in player_stats.items():
        if stats['MVP次数'] > 0:
            mvp_rate = stats['MVP次数'] / stats['总场次'] if stats['总场次'] > 0 else 0
            data.append({
                '玩家': player,
                'MVP次数': stats['MVP次数'],
                '总场次': stats['总场次'],
                'MVP率': f"{mvp_rate * 100:.1f}%"
            })
    
    df = pd.DataFrame(data)
    if len(df) > 0:
        df = df.sort_values(by='MVP次数', ascending=False).reset_index(drop=True)
        df.index = df.index + 1
    return df


def _create_hero_pool_leaderboard(player_stats: dict) -> pd.DataFrame:
    """生成英雄池排行榜"""
    data = []
    for player, stats in player_stats.items():
        if stats['总场次'] > 0:
            hero_pool_size = len(stats['英雄池'])
            data.append({
                '玩家': player,
                '英雄池数量': hero_pool_size,
                '总场次': stats['总场次'],
                '平均每英雄场次': round(stats['总场次'] / hero_pool_size, 2) if hero_pool_size > 0 else 0
            })
    
    df = pd.DataFrame(data)
    if len(df) > 0:
        df = df.sort_values(by=['英雄池数量', '总场次'], ascending=[False, False]).reset_index(drop=True)
        df.index = df.index + 1
    return df


def _create_streak_leaderboard(player_stats: dict) -> pd.DataFrame:
    """生成连胜排行榜"""
    data = [
        {'玩家': p, '最长连胜': s['最长连胜'], '当前连胜': s['连胜']}
        for p, s in player_stats.items() if s['最长连胜'] > 0
    ]
    
    df = pd.DataFrame(data)
    if len(df) > 0:
        df = df.sort_values(by=['最长连胜', '当前连胜'], ascending=[False, False]).reset_index(drop=True)
        df.index = df.index + 1
    return df


def _create_activity_leaderboard(player_stats: dict) -> pd.DataFrame:
    """生成活跃度排行榜"""
    data = []
    for player, stats in player_stats.items():
        if stats['总场次'] > 0:
            active_days = len(stats['每日比赛'])
            avg_games = stats['总场次'] / active_days if active_days > 0 else 0
            data.append({
                '玩家': player,
                '活跃天数': active_days,
                '总场次': stats['总场次'],
                '场均比赛': f"{avg_games:.1f}"
            })
    
    df = pd.DataFrame(data)
    if len(df) > 0:
        df = df.sort_values(by=['活跃天数', '总场次'], ascending=[False, False]).reset_index(drop=True)
        df.index = df.index + 1
    return df


def _create_teammate_leaderboard(player_pair_stats: dict, min_games: int = 2) -> pd.DataFrame:
    """生成最佳队友组合排行榜"""
    data = []
    for (p1, p2), stats in player_pair_stats.items():
        if stats['一起出场'] >= min_games:
            win_rate = stats['一起胜利'] / stats['一起出场']
            data.append({
                '队友组合': f"{p1} & {p2}",
                '一起出场': stats['一起出场'],
                '胜率': win_rate,
                '胜率百分比': f"{win_rate * 100:.1f}%"
            })
    
    df = pd.DataFrame(data)
    if len(df) > 0:
        df = df.sort_values(by=['胜率', '一起出场'], ascending=[False, False]).reset_index(drop=True)
        df.index = df.index + 1
    return df


def _create_hero_combo_leaderboard(hero_pair_stats: dict, min_games: int = 2) -> pd.DataFrame:
    """生成最佳英雄组合排行榜"""
    data = []
    for (h1, h2), stats in hero_pair_stats.items():
        if stats['一起出场'] >= min_games:
            win_rate = stats['一起胜利'] / stats['一起出场']
            data.append({
                '英雄组合': f"{h1} & {h2}",
                '一起出场': stats['一起出场'],
                '胜率': win_rate,
                '胜率百分比': f"{win_rate * 100:.1f}%"
            })
    
    df = pd.DataFrame(data)
    if len(df) > 0:
        df = df.sort_values(by=['胜率', '一起出场'], ascending=[False, False]).reset_index(drop=True)
        df.index = df.index + 1
    return df


def _create_daily_stats(df: pd.DataFrame) -> pd.DataFrame:
    """生成每日统计"""
    data = []
    for date in sorted(df['比赛时间'].unique()):
        day_df = df[df['比赛时间'] == date]
        blue_w = len(day_df[day_df['胜方'] == '蓝'])
        red_w = len(day_df[day_df['胜方'] == '红'])
        total = len(day_df)
        data.append({
            '日期': date,
            '比赛场次': total,
            '蓝方胜场': blue_w,
            '红方胜场': red_w,
            '蓝方胜率': f"{blue_w/total*100:.1f}%" if total > 0 else "0%",
            '红方胜率': f"{red_w/total*100:.1f}%" if total > 0 else "0%"
        })
    
    result_df = pd.DataFrame(data)
    if len(result_df) > 0:
        result_df.index = result_df.index + 1
    return result_df


def _create_position_leaderboards(player_stats: dict, hero_stats: dict) -> dict:
    """生成各分路排行榜"""
    result = {}
    
    for position in POSITIONS:
        # 玩家分路排行
        player_data = []
        for player, stats in player_stats.items():
            games = stats[f'{position}场次']
            if games >= 1:
                wins = stats[f'{position}胜场']
                win_rate = wins / games
                player_data.append({
                    '玩家': player,
                    '场次': games,
                    '胜场': wins,
                    '胜率': win_rate,
                    '胜率百分比': f"{win_rate * 100:.1f}%"
                })
        
        player_df = pd.DataFrame(player_data)
        if len(player_df) > 0:
            player_df = player_df.sort_values(by=['胜率', '场次'], ascending=[False, False]).reset_index(drop=True)
            player_df.index = player_df.index + 1
        
        # 英雄分路排行
        hero_data = []
        for hero, stats in hero_stats.items():
            games = stats[f'{position}场次']
            if games >= 1:
                wins = stats[f'{position}胜场']
                win_rate = wins / games
                hero_data.append({
                    '英雄': hero,
                    '场次': games,
                    '胜场': wins,
                    '胜率': win_rate,
                    '胜率百分比': f"{win_rate * 100:.1f}%"
                })
        
        hero_df = pd.DataFrame(hero_data)
        if len(hero_df) > 0:
            hero_df = hero_df.sort_values(by=['胜率', '场次'], ascending=[False, False]).reset_index(drop=True)
            hero_df.index = hero_df.index + 1
        
        result[position] = {'player': player_df, 'hero': hero_df}
    
    return result


def _create_hero_player_leaderboard(hero_stats: dict) -> dict:
    """生成英雄-玩家胜率榜"""
    result = {}
    for hero, stats in hero_stats.items():
        player_list = []
        for player, games in stats['玩家场次'].items():
            if games >= 1:
                wins = stats['玩家胜场'].get(player, 0)
                win_rate = wins / games
                player_list.append({
                    '玩家': player,
                    '场次': games,
                    '胜场': wins,
                    '胜率': win_rate,
                    '胜率百分比': f"{win_rate * 100:.1f}%"
                })
        player_list.sort(key=lambda x: (x['胜率'], x['场次']), reverse=True)
        result[hero] = player_list
    return result


def _create_player_hero_leaderboard(player_stats: dict) -> dict:
    """生成玩家-英雄胜率榜"""
    result = {}
    for player, stats in player_stats.items():
        hero_list = []
        for hero, games in stats['英雄场次'].items():
            if games >= 1:
                wins = stats['英雄胜场'].get(hero, 0)
                win_rate = wins / games
                hero_list.append({
                    '英雄': hero,
                    '场次': games,
                    '胜场': wins,
                    '胜率': win_rate,
                    '胜率百分比': f"{win_rate * 100:.1f}%"
                })
        hero_list.sort(key=lambda x: (x['胜率'], x['场次']), reverse=True)
        result[player] = hero_list
    return result


def _create_bounty_leaderboard(df: pd.DataFrame) -> dict:
    """
    生成赏金榜（增强版）
    规则：胜方全员+2元，胜方MVP+1元，败方MVP+1元
    
    返回:
        dict: 包含以下内容
            - leaderboard: 总排行榜DataFrame（含每日明细列）
            - daily_detail: 每日赏金明细 {日期: {玩家: 金额}}
            - pool_info: 奖金池信息 {初始值, 已发放, 剩余}
            - total_distributed: 总发放金额
    """
    # 2026年奖金池初始值
    POOL_INITIAL = 401.5
    
    # 总赏金统计
    bounty_total = defaultdict(float)
    # 每日赏金明细 {日期: {玩家: 金额}}
    bounty_daily = defaultdict(lambda: defaultdict(float))
    # 获取所有日期
    all_dates = sorted(df['比赛时间'].unique().tolist())
    
    for idx, row in df.iterrows():
        winner = row['胜方']
        match_date = row['比赛时间']
        
        # 获取蓝方玩家
        blue_players = []
        for pos in ['蓝方边', '蓝方野', '蓝方中', '蓝方射', '蓝方辅']:
            cell = row[pos]
            if isinstance(cell, str) and '-' in cell:
                player, _ = cell.split('-', 1)
                blue_players.append(player)
        
        # 获取红方玩家
        red_players = []
        for pos in ['红方边', '红方野', '红方中', '红方射', '红方辅']:
            cell = row[pos]
            if isinstance(cell, str) and '-' in cell:
                player, _ = cell.split('-', 1)
                red_players.append(player)
        
        # 胜方全员+2元
        win_players = blue_players if winner == '蓝' else red_players
        lose_players = red_players if winner == '蓝' else blue_players
        for p in win_players:
            bounty_total[p] += 2
            bounty_daily[match_date][p] += 2
        
        # 胜方MVP+1元
        win_mvp_pos = row['蓝方MVP'] if winner == '蓝' else row['红方MVP']
        if win_mvp_pos in MVP_INDEX_MAP:
            mvp_idx = MVP_INDEX_MAP[win_mvp_pos]
            if mvp_idx < len(win_players):
                bounty_total[win_players[mvp_idx]] += 1
                bounty_daily[match_date][win_players[mvp_idx]] += 1
        
        # 败方MVP+1元
        lose_mvp_pos = row['红方MVP'] if winner == '蓝' else row['蓝方MVP']
        if lose_mvp_pos in MVP_INDEX_MAP:
            mvp_idx = MVP_INDEX_MAP[lose_mvp_pos]
            if mvp_idx < len(lose_players):
                bounty_total[lose_players[mvp_idx]] += 1
                bounty_daily[match_date][lose_players[mvp_idx]] += 1
    
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
    result_df = pd.DataFrame(leaderboard_data)
    if len(result_df) > 0:
        result_df = result_df.sort_values(by='总赏金', ascending=False).reset_index(drop=True)
        result_df.index = result_df.index + 1
    
    # 奖金池信息
    pool_info = {
        'initial': POOL_INITIAL,
        'distributed': total_distributed,
        'remaining': POOL_INITIAL - total_distributed
    }
    
    return {
        'leaderboard': result_df,
        'daily_detail': dict(bounty_daily),
        'pool_info': pool_info,
        'total_distributed': total_distributed,
        'dates': all_dates
    }


# ========== 数据加载函数 ==========
def load_match_data(file_path: str) -> pd.DataFrame:
    """
    加载比赛数据
    
    参数:
        file_path: Excel文件路径
    
    返回:
        pd.DataFrame: 包含比赛数据的DataFrame
    """
    data = pd.read_excel(file_path)
    data['比赛时间'] = data['比赛时间'].apply(
        lambda x: x.strftime("%Y-%m-%d") if hasattr(x, 'strftime') else str(x)
    )
    return data[COLUMNS].copy()


def filter_by_date_range(df: pd.DataFrame, start_date: str = None, end_date: str = None) -> pd.DataFrame:
    """
    按日期范围筛选数据
    
    参数:
        df: 原始DataFrame
        start_date: 开始日期 (YYYY-MM-DD)
        end_date: 结束日期 (YYYY-MM-DD)
    
    返回:
        pd.DataFrame: 筛选后的DataFrame
    """
    result = df.copy()
    if start_date:
        result = result[result['比赛时间'] >= start_date]
    if end_date:
        result = result[result['比赛时间'] <= end_date]
    return result


def filter_by_year(df: pd.DataFrame, year: int) -> pd.DataFrame:
    """
    按年份筛选数据
    
    参数:
        df: 原始DataFrame
        year: 年份 (如 2025, 2026)
    
    返回:
        pd.DataFrame: 筛选后的DataFrame
    """
    return df[df['比赛时间'].str.startswith(str(year))].copy()


# ========== 测试代码 ==========
if __name__ == "__main__":
    # 测试加载和计算
    file_path = 'C:/Files/Ubiquant/code/HOK/hok_bp/practicing/内战data/内战计分表 - 2026.xlsx'
    
    print("正在加载数据...")
    df = load_match_data(file_path)
    print(f"总共加载 {len(df)} 场比赛")
    
    print("\n按年份筛选:")
    df_2025 = filter_by_year(df, 2025)
    df_2026 = filter_by_year(df, 2026)
    print(f"2025年: {len(df_2025)} 场比赛")
    print(f"2026年: {len(df_2026)} 场比赛")
    
    print("\n正在计算统计数据...")
    stats = calculate_all_stats(df)
    
    print(f"\n基础统计:")
    for key, value in stats['basic_stats'].items():
        print(f"  {key}: {value}")
    
    print(f"\n玩家排行榜 TOP 5:")
    print(stats['player_leaderboard'].head())
    
    print(f"\n英雄排行榜 TOP 5:")
    print(stats['hero_leaderboard'].head())
    
    print(f"\n赏金榜 TOP 5:")
    bounty_data = stats['bounty_leaderboard']
    print(bounty_data['leaderboard'].head())
    print(f"\n奖金池信息:")
    print(f"  初始值: {bounty_data['pool_info']['initial']}元")
    print(f"  已发放: {bounty_data['pool_info']['distributed']}元")
    print(f"  剩余: {bounty_data['pool_info']['remaining']}元")
    
    print("\n计算完成！")
