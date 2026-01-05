"""
王者荣耀内战数据分析系统 - Flask后端API

提供RESTful API接口，支持前端实时筛选和数据查询
"""

from flask import Flask, jsonify, request, render_template, send_from_directory
from flask_cors import CORS
import pandas as pd
import os
from datetime import datetime

# 导入计算模块
from match_analyzer import (
    load_match_data, 
    calculate_all_stats, 
    filter_by_date_range, 
    filter_by_year
)

app = Flask(__name__, static_folder='static', template_folder='templates')
CORS(app)

# ========== 全局数据 ==========
DATA_FILE = 'C:/Files/Ubiquant/code/HOK/hok_bp/practicing/内战data/内战计分表 - 2026.xlsx'
_cached_df = None
_cached_stats = {}


def get_base_data():
    """获取基础数据（带缓存）"""
    global _cached_df
    if _cached_df is None:
        _cached_df = load_match_data(DATA_FILE)
    return _cached_df


def get_stats_for_filter(start_date=None, end_date=None, year=None):
    """
    根据筛选条件获取统计数据
    
    参数:
        start_date: 开始日期
        end_date: 结束日期
        year: 年份（如果指定，优先使用年份筛选）
    """
    df = get_base_data()
    
    # 构建缓存键
    cache_key = f"{start_date}_{end_date}_{year}"
    
    if cache_key in _cached_stats:
        return _cached_stats[cache_key]
    
    # 筛选数据
    if year:
        filtered_df = filter_by_year(df, year)
    else:
        filtered_df = filter_by_date_range(df, start_date, end_date)
    
    # 计算统计
    stats = calculate_all_stats(filtered_df)
    
    # 缓存结果
    _cached_stats[cache_key] = stats
    
    return stats


def df_to_json(df):
    """将DataFrame转换为JSON格式"""
    if df is None or len(df) == 0:
        return []
    return df.reset_index().to_dict(orient='records')


# ========== API路由 ==========

@app.route('/')
def index():
    """主页 - 返回交互式报告页面"""
    return render_template('report.html')


@app.route('/api/dates')
def get_available_dates():
    """获取可用的日期范围"""
    df = get_base_data()
    dates = sorted(df['比赛时间'].unique().tolist())
    years = sorted(list(set(d[:4] for d in dates)))
    
    return jsonify({
        'success': True,
        'data': {
            'dates': dates,
            'years': years,
            'min_date': dates[0] if dates else None,
            'max_date': dates[-1] if dates else None,
            'total_matches': len(df)
        }
    })


@app.route('/api/stats')
def get_stats():
    """
    获取统计数据
    
    查询参数:
        - start_date: 开始日期 (YYYY-MM-DD)
        - end_date: 结束日期 (YYYY-MM-DD)
        - year: 年份 (2025, 2026, all)
    """
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    year = request.args.get('year')
    
    # 处理年份参数
    if year and year.lower() != 'all':
        try:
            year = int(year)
        except ValueError:
            year = None
    else:
        year = None
    
    try:
        stats = get_stats_for_filter(start_date, end_date, year)
        
        if stats is None:
            return jsonify({
                'success': True,
                'data': {
                    'has_data': False,
                    'message': '所选时间范围内没有比赛数据'
                }
            })
        
        return jsonify({
            'success': True,
            'data': {
                'has_data': True,
                'basic_stats': stats['basic_stats'],
                'player_leaderboard': df_to_json(stats['player_leaderboard'].head(20)),
                'hero_leaderboard': df_to_json(stats['hero_leaderboard'].head(20)),
                'mvp_leaderboard': df_to_json(stats['mvp_leaderboard'].head(10)),
                'hero_pool_leaderboard': df_to_json(stats['hero_pool_leaderboard'].head(10)),
                'streak_leaderboard': df_to_json(stats['streak_leaderboard'].head(10)),
                'activity_leaderboard': df_to_json(stats['activity_leaderboard'].head(10)),
                'teammate_leaderboard': df_to_json(stats['teammate_leaderboard'].head(10)),
                'hero_combo_leaderboard': df_to_json(stats['hero_combo_leaderboard'].head(10)),
                'daily_stats': df_to_json(stats['daily_stats']),
                'bounty_leaderboard': df_to_json(stats['bounty_leaderboard']['leaderboard'].head(20)),
                'bounty_pool_info': stats['bounty_leaderboard']['pool_info'],
                'bounty_dates': stats['bounty_leaderboard']['dates'],
            }
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@app.route('/api/stats/position/<position>')
def get_position_stats(position):
    """获取分路统计数据"""
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    year = request.args.get('year')
    
    if year and year.lower() != 'all':
        try:
            year = int(year)
        except ValueError:
            year = None
    else:
        year = None
    
    try:
        stats = get_stats_for_filter(start_date, end_date, year)
        
        if stats is None:
            return jsonify({'success': True, 'data': {'has_data': False}})
        
        pos_stats = stats['position_leaderboards'].get(position, {})
        
        return jsonify({
            'success': True,
            'data': {
                'has_data': True,
                'player_leaderboard': df_to_json(pos_stats.get('player', pd.DataFrame()).head(10)),
                'hero_leaderboard': df_to_json(pos_stats.get('hero', pd.DataFrame()).head(10))
            }
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/stats/player/<player>')
def get_player_detail(player):
    """获取玩家详细数据"""
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    year = request.args.get('year')
    
    if year and year.lower() != 'all':
        try:
            year = int(year)
        except ValueError:
            year = None
    else:
        year = None
    
    try:
        stats = get_stats_for_filter(start_date, end_date, year)
        
        if stats is None or player not in stats['player_stats']:
            return jsonify({'success': True, 'data': {'has_data': False}})
        
        player_data = stats['player_stats'][player]
        hero_list = stats['player_hero_leaderboard'].get(player, [])
        
        return jsonify({
            'success': True,
            'data': {
                'has_data': True,
                'player': player,
                'stats': {
                    '总场次': player_data['总场次'],
                    '总胜场': player_data['总胜场'],
                    '总胜率': f"{player_data['总胜场']/player_data['总场次']*100:.1f}%" if player_data['总场次'] > 0 else "0%",
                    'MVP次数': player_data['MVP次数'],
                    '英雄池数量': len(player_data['英雄池']),
                    '最长连胜': player_data['最长连胜'],
                    '当前连胜': player_data['连胜'],
                    '边路场次': player_data['边路场次'],
                    '打野场次': player_data['打野场次'],
                    '中路场次': player_data['中路场次'],
                    '发育路场次': player_data['发育路场次'],
                    '游走场次': player_data['游走场次'],
                },
                'hero_list': hero_list[:10]
            }
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/stats/hero/<hero>')
def get_hero_detail(hero):
    """获取英雄详细数据"""
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    year = request.args.get('year')
    
    if year and year.lower() != 'all':
        try:
            year = int(year)
        except ValueError:
            year = None
    else:
        year = None
    
    try:
        stats = get_stats_for_filter(start_date, end_date, year)
        
        if stats is None or hero not in stats['hero_stats']:
            return jsonify({'success': True, 'data': {'has_data': False}})
        
        hero_data = stats['hero_stats'][hero]
        player_list = stats['hero_player_leaderboard'].get(hero, [])
        
        return jsonify({
            'success': True,
            'data': {
                'has_data': True,
                'hero': hero,
                'stats': {
                    '总场次': hero_data['总场次'],
                    '总胜场': hero_data['总胜场'],
                    '总胜率': f"{hero_data['总胜场']/hero_data['总场次']*100:.1f}%" if hero_data['总场次'] > 0 else "0%",
                    '使用玩家数': len(hero_data['玩家场次']),
                    '边路场次': hero_data['边路场次'],
                    '打野场次': hero_data['打野场次'],
                    '中路场次': hero_data['中路场次'],
                    '发育路场次': hero_data['发育路场次'],
                    '游走场次': hero_data['游走场次'],
                },
                'player_list': player_list[:10]
            }
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/matches')
def get_matches():
    """获取比赛记录"""
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    year = request.args.get('year')
    page = int(request.args.get('page', 1))
    per_page = int(request.args.get('per_page', 20))
    
    df = get_base_data()
    
    # 筛选
    if year and year.lower() != 'all':
        try:
            df = filter_by_year(df, int(year))
        except ValueError:
            pass
    else:
        df = filter_by_date_range(df, start_date, end_date)
    
    # 分页
    total = len(df)
    start_idx = (page - 1) * per_page
    end_idx = start_idx + per_page
    
    matches = df.iloc[start_idx:end_idx].to_dict(orient='records')
    
    return jsonify({
        'success': True,
        'data': {
            'matches': matches,
            'total': total,
            'page': page,
            'per_page': per_page,
            'total_pages': (total + per_page - 1) // per_page
        }
    })


@app.route('/api/refresh')
def refresh_data():
    """刷新数据（清除缓存）"""
    global _cached_df, _cached_stats
    _cached_df = None
    _cached_stats = {}
    
    return jsonify({
        'success': True,
        'message': '数据已刷新'
    })


# ========== 静态文件路由 ==========
@app.route('/static/<path:filename>')
def serve_static(filename):
    return send_from_directory(app.static_folder, filename)


# ========== 主程序 ==========
if __name__ == '__main__':
    # 确保模板目录存在
    os.makedirs('templates', exist_ok=True)
    os.makedirs('static', exist_ok=True)
    
    print("=" * 60)
    print("🏆 王者荣耀内战数据分析系统")
    print("=" * 60)
    print(f"\n服务器启动中...")
    print(f"访问地址: http://127.0.0.1:5000")
    print(f"\nAPI接口:")
    print(f"  GET /api/dates        - 获取可用日期范围")
    print(f"  GET /api/stats        - 获取统计数据 (支持 start_date, end_date, year 参数)")
    print(f"  GET /api/matches      - 获取比赛记录 (支持分页)")
    print(f"  GET /api/refresh      - 刷新数据缓存")
    print("=" * 60)
    
    app.run(debug=True, host='0.0.0.0', port=5000)
