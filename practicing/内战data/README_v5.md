# 王者荣耀内战数据分析系统 v5.0

## 🎯 项目结构

本系统将**计算逻辑**与**展示逻辑**完全解耦，提供了一个可扩展的交互式数据分析平台。

```
内战data/
├── match_analyzer.py     # 核心计算模块（纯Python，无UI依赖）
├── app.py               # Flask后端API服务
├── templates/
│   └── report.html      # 交互式前端页面
├── requirements.txt     # Python依赖
├── 启动服务.bat          # Windows一键启动脚本
└── cal_rate_report.py   # 旧版静态报告生成器（保留）
```

## 🏗️ 架构说明

### 1. 核心计算模块 (`match_analyzer.py`)

**核心函数**：`calculate_all_stats(df)` 

接收一个比赛数据的DataFrame，返回一个包含所有统计结果的字典：

```python
from match_analyzer import load_match_data, calculate_all_stats, filter_by_date_range

# 加载数据
df = load_match_data('内战计分表.xlsx')

# 按日期范围筛选
df_filtered = filter_by_date_range(df, start_date='2025-12-01', end_date='2025-12-31')

# 计算所有统计数据
stats = calculate_all_stats(df_filtered)

# 访问各种统计结果
print(stats['basic_stats'])           # 基础统计
print(stats['player_leaderboard'])    # 玩家排行榜
print(stats['hero_leaderboard'])      # 英雄排行榜
print(stats['bounty_leaderboard'])    # 赏金榜
# ... 更多数据表
```

**返回的数据结构**：
```python
{
    'df': DataFrame,                    # 原始数据
    'basic_stats': dict,                # 基础统计（比赛数、玩家数等）
    'player_stats': dict,               # 玩家详细统计
    'hero_stats': dict,                 # 英雄详细统计
    'player_leaderboard': DataFrame,    # 玩家综合排行榜
    'hero_leaderboard': DataFrame,      # 英雄胜率排行榜
    'mvp_leaderboard': DataFrame,       # MVP排行榜
    'hero_pool_leaderboard': DataFrame, # 英雄池排行榜
    'streak_leaderboard': DataFrame,    # 连胜排行榜
    'activity_leaderboard': DataFrame,  # 活跃度排行榜
    'teammate_leaderboard': DataFrame,  # 最佳队友组合排行榜
    'hero_combo_leaderboard': DataFrame,# 最佳英雄组合排行榜
    'daily_stats': DataFrame,           # 每日统计
    'position_leaderboards': dict,      # 各分路排行榜
    'hero_player_leaderboard': dict,    # 英雄-玩家胜率榜
    'player_hero_leaderboard': dict,    # 玩家-英雄胜率榜
    'bounty_leaderboard': DataFrame,    # 赏金榜
}
```

### 2. Flask API服务 (`app.py`)

提供RESTful API接口，支持前端实时查询：

| API端点 | 方法 | 说明 |
|---------|------|------|
| `/` | GET | 主页（交互式报告页面） |
| `/api/dates` | GET | 获取可用的日期范围 |
| `/api/stats?year=2025` | GET | 获取统计数据（支持year, start_date, end_date参数） |
| `/api/stats/position/<pos>` | GET | 获取分路统计 |
| `/api/stats/player/<name>` | GET | 获取玩家详细数据 |
| `/api/stats/hero/<name>` | GET | 获取英雄详细数据 |
| `/api/matches?page=1` | GET | 获取比赛记录（分页） |
| `/api/refresh` | GET | 刷新数据缓存 |

### 3. 交互式前端 (`templates/report.html`)

- 支持**年份快速筛选**（2025年、2026年、全部）
- 支持**自定义日期范围**筛选
- **实时查询**：筛选条件变化后，调用后端API重新计算并展示
- 多Tab页面：总览、赏金榜、玩家排行、英雄数据、组合数据、比赛记录

## 🚀 快速开始

### 方法1：双击启动（推荐）

直接双击 `启动服务.bat`，然后在浏览器打开 http://127.0.0.1:5000

### 方法2：命令行启动

```bash
# 安装依赖
pip install -r requirements.txt

# 启动服务
python app.py

# 访问 http://127.0.0.1:5000
```

### 方法3：仅使用计算模块（不需要Web服务）

```python
from match_analyzer import load_match_data, calculate_all_stats

df = load_match_data('你的数据文件.xlsx')
stats = calculate_all_stats(df)

# 导出为Excel
stats['player_leaderboard'].to_excel('玩家排行榜.xlsx')
stats['hero_leaderboard'].to_excel('英雄排行榜.xlsx')
```

## 📊 数据格式要求

Excel文件需要包含以下列：
- `比赛ID` - 比赛唯一标识
- `比赛时间` - 日期格式
- `胜方` - "蓝" 或 "红"
- `蓝方边`, `蓝方野`, `蓝方中`, `蓝方射`, `蓝方辅` - 格式为 "玩家名-英雄名"
- `红方边`, `红方野`, `红方中`, `红方射`, `红方辅` - 格式为 "玩家名-英雄名"
- `蓝方MVP`, `红方MVP` - "边"/"野"/"中"/"射"/"辅"

## 🔧 扩展开发

### 添加新的统计指标

在 `match_analyzer.py` 中添加新的计算函数：

```python
def _create_my_custom_leaderboard(player_stats: dict) -> pd.DataFrame:
    """自定义排行榜"""
    data = []
    for player, stats in player_stats.items():
        # 你的计算逻辑
        data.append({...})
    return pd.DataFrame(data)
```

然后在 `calculate_all_stats()` 函数的返回值中添加：
```python
result = {
    ...
    'my_custom_leaderboard': _create_my_custom_leaderboard(player_stats),
}
```

### 添加新的API端点

在 `app.py` 中添加：

```python
@app.route('/api/my-custom-data')
def get_my_custom_data():
    stats = get_stats_for_filter(...)
    return jsonify({
        'success': True,
        'data': df_to_json(stats['my_custom_leaderboard'])
    })
```

### 添加前端展示

在 `templates/report.html` 中添加新的Tab和渲染逻辑。

## 📝 更新日志

### v5.0 (2025-01-05)
- 🏗️ **架构重构**：将计算逻辑与展示逻辑完全解耦
- 🌐 **交互式前端**：支持实时筛选时间范围
- 🔌 **RESTful API**：提供标准化的数据接口
- 📦 **模块化设计**：核心计算模块可独立使用

### v4.0 (之前版本)
- 静态HTML报告生成
- 年份分页功能
- 赏金榜计算

## 🤝 贡献者

- Yuanhang Zhang

---
*Copyright: Yuanhang Zhang -- v5.0 (Interactive)*
