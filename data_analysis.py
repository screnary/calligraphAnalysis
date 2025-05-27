# 基础包
import os
import re
from datetime import datetime
import argparse

# 数据处理包
import pandas as pd
import numpy as np

import process_excel as pe

# Excel文件处理包
import openpyxl  # 处理.xlsx文件
import xlrd      # 读取.xls文件（旧版Excel）
import xlwt      # 写入.xls文件（旧版Excel）
import xlsxwriter # 另一个写入.xlsx的库，功能更强大

# 数据可视化（如果需要在Excel中创建图表）
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.express as px

# 其他实用包
import logging
from typing import List, Dict, Optional, Union

import pdb

# 设置pandas显示选项
pd.set_option('display.max_columns', None)  # 显示所有列
pd.set_option('display.max_rows', 100)      # 显示最多100行
pd.set_option('display.width', None)        # 自动调整显示宽度
pd.set_option('display.max_colwidth', None) # 不限制列宽

# 设置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# 忽略警告（可选）
import warnings
warnings.filterwarnings('ignore')


# arg parser
def parse_arguments():
    """解析命令行参数"""
    parser = argparse.ArgumentParser(
        description='Excel时间数据分析工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  %(prog)s -f data/example.xlsx -c 书体
  %(prog)s -f data/example.xlsx -c 书体 -t 刻石时间 -s 400 -e 1000 -p 50
  %(prog)s -f data/example.xlsx -c 书体 --output results.xlsx --plot
  %(prog)s --list-files data/
  %(prog)s -f data/example.xlsx --info
        """
    )
    
    # 必需参数
    parser.add_argument('-f', '--file', 
                        type=str,
                        help='Excel文件路径')
    
    parser.add_argument('-c', '--category', 
                        type=str,
                        help='类别列名称（如：书体）')
    
    # 可选参数
    parser.add_argument('-t', '--time', 
                        type=str,
                        default='刻石时间',
                        help='时间列名称（默认：刻石时间）')
    
    parser.add_argument('-s', '--start', 
                        type=int,
                        default=400,
                        help='起始年份（默认：400）')
    
    parser.add_argument('-e', '--end', 
                        type=int,
                        default=1000,
                        help='结束年份（默认：1000）')
    
    parser.add_argument('-p', '--period', 
                        type=int,
                        default=50,
                        help='时间段长度（默认：50年）')
    
    parser.add_argument('--sheet', 
                        type=str,
                        default=None,
                        help='工作表名称或索引（默认：第一个工作表）')
    
    parser.add_argument('-o', '--output', 
                        type=str,
                        default=None,
                        help='输出文件路径（保存分析结果）')
    
    parser.add_argument('--plot', 
                        action='store_true',
                        help='生成可视化图表')
    
    parser.add_argument('--save-plot', 
                        type=str,
                        default=None,
                        help='保存图表到文件')
    
    # 信息查询参数
    parser.add_argument('--info', 
                        action='store_true',
                        help='显示文件信息（工作表、列名等）')
    
    parser.add_argument('--list-files', 
                        type=str,
                        default=None,
                        help='列出目录中的所有Excel文件')
    
    parser.add_argument('-v', '--verbose', 
                        action='store_true',
                        help='显示详细信息')
    
    return parser.parse_args()


# conduct analyze
def analyze_excel_with_args(args):
    """根据命令行参数分析Excel文件"""
    try:
        # 读取Excel文件
        if args.sheet is not None:
            # 尝试将sheet参数转换为整数（索引）
            try:
                sheet = int(args.sheet)
            except ValueError:
                sheet = args.sheet
        else:
            sheet = 0  # 默认第一个工作表
        
        print(f"\n读取文件: {args.file}")
        # df = pd.read_excel(args.file, sheet_name=sheet)
        df = pe.read_excel_advanced(args.file, sheet_name=sheet)
        
        # 获取实际的工作表名称
        with pd.ExcelFile(args.file) as xls:
            actual_sheet_name = xls.sheet_names[sheet] if isinstance(sheet, int) else sheet
            print(f"使用工作表: '{actual_sheet_name}'")
        
        print(f"数据形状: {df.shape[0]} 行 × {df.shape[1]} 列")
        
        # 检查列是否存在
        if args.time not in df.columns:
            print(f"\n错误: 时间列 '{args.time}' 不存在")
            print(f"可用的列: {list(df.columns)}")
            return None
        
        if args.category not in df.columns:
            print(f"\n错误: 类别列 '{args.category}' 不存在")
            print(f"可用的列: {list(df.columns)}")
            return None
        
        # 执行分析
        print(f"\n执行时间分组统计...")
        print(f"- 时间列: {args.time}")
        print(f"- 类别列: {args.category}")
        print(f"- 时间范围: {args.start} - {args.end}")
        print(f"- 时间段长度: {args.period} 年")
        
        result = group_by_time_period(
            df=df,
            time_column=args.time,
            category_column=args.category,
            start_year=args.start,
            end_year=args.end,
            period=args.period,
            show_details=args.verbose
        )
        
        # 显示结果
        print("\n统计结果:")
        print("=" * 80)
        print(result)
        
        # 保存结果
        if args.output:
            save_results(result, args.output, df, args)
            print(f"\n结果已保存到: {args.output}")
        
        # 生成图表
        if args.plot or args.save_plot:
            plot_results_plotly(result, args)
        
        return result
        
    except Exception as e:
        print(f"\n错误: {str(e)}")
        return None

def save_results(result: pd.DataFrame, output_path: str, 
                 original_df: pd.DataFrame, args):
    """保存分析结果到Excel文件"""
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # 保存统计结果
        result.to_excel(writer, sheet_name='时间统计结果')
        
        # 保存参数信息
        params_df = pd.DataFrame({
            '参数': ['文件路径', '时间列', '类别列', '起始年份', 
                    '结束年份', '时间段长度', '原始数据行数'],
            '值': [args.file, args.time, args.category, args.start, 
                   args.end, args.period, len(original_df)]
        })
        params_df.to_excel(writer, sheet_name='分析参数', index=False)
        
        # 保存数据摘要
        summary_df = pd.DataFrame({
            '统计项': ['总记录数', '有效时间记录数', '类别总数'],
            '数量': [
                len(original_df),
                original_df[args.time].notna().sum(),
                original_df[args.category].nunique()
            ]
        })
        summary_df.to_excel(writer, sheet_name='数据摘要', index=False)


def plot_results_plotly(result: pd.DataFrame, args):
    """使用Plotly生成可视化图表"""
    # 准备数据（排除总计行和列）
    plot_data = result.iloc[:-1, :-1]
    
    # 创建子图布局
    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=(
            f'{args.category}随时间变化（堆叠图）',
            f'{args.category}时间分布热力图',
            f'所有{args.category}的时间趋势',
            f'{args.category}总体分布'
        ),
        specs=[
            [{"type": "bar"}, {"type": "heatmap"}],
            [{"type": "scatter"}, {"type": "pie"}]
        ],
        horizontal_spacing=0.20,
        vertical_spacing=0.15
    )
    
    # 1. 堆叠条形图
    for col in plot_data.columns:
        fig.add_trace(
            go.Bar(
                name=col,
                x=plot_data.index,
                y=plot_data[col],
                text=plot_data[col],
                textposition='auto',
            ),
            row=1, col=1
        )
    
    # 更新堆叠条形图布局
    fig.update_xaxes(title_text="时间段", row=1, col=1)
    fig.update_yaxes(title_text="数量", row=1, col=1)
    fig.update_layout(barmode='stack')
    
    # 2. 热力图
    # 准备热力图数据
    heatmap_data = plot_data.T.values
    heatmap_text = [[str(val) for val in row] for row in heatmap_data]
    
    fig.add_trace(
        go.Heatmap(
            z=heatmap_data,
            x=plot_data.index.tolist(),
            y=plot_data.columns.tolist(),
            text=heatmap_text,
            texttemplate="%{text}",
            colorscale='YlOrRd',
            showscale=True,
            colorbar=dict(
                x=1.12,  # 将colorbar移到更右边
                xpad=10,
                len=0.45,  # 缩短colorbar长度
                y=0.77,
                yanchor='top'
            ),
            showlegend=False  # 热力图不需要在图例中显示
        ),
        row=1, col=2
    )
    
    fig.update_xaxes(title_text="时间段", row=1, col=2)
    fig.update_yaxes(title_text=args.category, row=1, col=2)
    
    # 3. 折线图（所有类别）
    show_in_legend = len(plot_data.columns) <= 10
    for col in plot_data.columns:
        fig.add_trace(
            go.Scatter(
                name=col,
                x=plot_data.index,
                y=plot_data[col],
                mode='lines+markers',
                showlegend=show_in_legend,
                legendgroup='lines'  # 将折线图的图例分组
            ),
            row=2, col=1
        )
    
    fig.update_xaxes(title_text="时间段", row=2, col=1)
    fig.update_yaxes(title_text="数量", row=2, col=1)
    
    # 如果类别太多，添加说明
    if len(plot_data.columns) > 10:
        fig.add_annotation(
            text=f"共{len(plot_data.columns)}个类别（图例已隐藏）",
            xref="paper", yref="paper",
            x=0.25, y=-0.1,
            showarrow=False,
            font=dict(size=10)
        )
    
    # 4. 饼图（总体分布）
    category_totals = plot_data.sum().sort_values(ascending=False)
    
    # 如果类别太多，只显示前10个
    if len(category_totals) > 15:
        top_10 = category_totals.head(10)
        others = category_totals.iloc[10:].sum()
        if others > 0:
            display_data = pd.concat([top_10, pd.Series({'其他': others})])
        else:
            display_data = top_10
    else:
        display_data = category_totals
    
    fig.add_trace(
        go.Pie(
            labels=display_data.index,
            values=display_data.values,
            textinfo='label+percent',
            hovertemplate='%{label}: %{value}<br>占比: %{percent}<extra></extra>',
            showlegend=False  # 饼图不需要在图例中显示
        ),
        row=2, col=2
    )
    
    # 更新整体布局
    fig.update_layout(
        title={
            'text': f'文件: {os.path.basename(args.file)} - {args.category}时间分析',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 20}
        },
        height=900,  # 增加高度
        width=1400,  # 设置宽度
        showlegend=True,
        barmode='stack',
        # 调整图例位置和样式
        legend=dict(
            yanchor="top",
            y=0.98,
            xanchor="left",
            x=1.02,  # 将图例放在图表右侧
            bgcolor="rgba(255, 255, 255, 0.8)",
            bordercolor="rgba(0, 0, 0, 0.2)",
            borderwidth=1,
            font=dict(size=10),
            itemsizing='constant',
            tracegroupgap=5
        ),
        # 调整边距
        margin=dict(
            l=50,
            r=250,  # 增加右边距给图例和colorbar留空间
            t=100,
            b=50
        ),
        font=dict(size=12)
    )
    
    # 保存或显示
    if args.save_plot:
        current_time = datetime.now()
        # 保存为HTML（交互式）
        html_file = args.save_plot.replace('.html', f'_{current_time}.html')
        # pdb.set_trace()
        fig.write_html(html_file)
        print(f"交互式图表已保存到: {html_file}")
    else:
        fig.show()


# Calculate stocastic info
def parse_time_value(time_str: str) -> Optional[int]:
    """
    解析时间字符串，提取数值
    
    参数:
        time_str: 时间字符串，如 "740左右", "650-655", "850"
    
    返回:
        解析后的整数值，无法解析返回None
    """
    if pd.isna(time_str) or time_str == '':
        return None
    
    # 转换为字符串并去除空格
    time_str = str(time_str).strip()
    
    # 情况1: 纯数字
    if time_str.isdigit():
        return int(time_str)
    
    # 情况2: xxx左右
    pattern_zuoyou = re.match(r'(\d+)\s*左右', time_str)
    if pattern_zuoyou:
        return int(pattern_zuoyou.group(1))
    
    # 情况3: xxx-xxx (取第一个数字)
    pattern_range = re.match(r'(\d+)\s*[-—]\s*\d+', time_str)
    if pattern_range:
        return int(pattern_range.group(1))
    
    # 情况4: 尝试提取任何数字
    numbers = re.findall(r'\d+', time_str)
    if numbers:
        return int(numbers[0])
    
    return None


def group_by_time_period(
    df: pd.DataFrame,
    time_column: str,
    category_column: str,
    start_year: int = 400,
    end_year: int = 1000,
    period: int = 50,
    show_details: bool = True
) -> pd.DataFrame:
    """
    按时间段分组统计各类别数量
    
    参数:
        df: 数据框
        time_column: 时间列名称（如"刻石时间"）
        category_column: 类别列名称（如"书体"）
        start_year: 起始年份
        end_year: 结束年份
        period: 时间段长度（默认50年）
        show_details: 是否显示详细信息
    
    返回:
        统计结果DataFrame
    """
    # 检查列是否存在
    if time_column not in df.columns:
        raise ValueError(f"列 '{time_column}' 不存在于数据中")
    if category_column not in df.columns:
        raise ValueError(f"列 '{category_column}' 不存在于数据中")
    
    # 复制数据避免修改原始数据
    data = df.copy()
    
    # 解析时间列
    print(f"正在解析 '{time_column}' 列...")
    data['parsed_time'] = data[time_column].apply(parse_time_value)
    
    # 显示解析情况
    if show_details:
        total_rows = len(data)
        parsed_rows = data['parsed_time'].notna().sum()
        print(f"总记录数: {total_rows}")
        print(f"成功解析: {parsed_rows} ({parsed_rows/total_rows*100:.1f}%)")
        print(f"无法解析: {total_rows - parsed_rows}")
        
        # 显示无法解析的样例
        unparsed = data[data['parsed_time'].isna()][time_column].dropna().unique()[:5]
        if len(unparsed) > 0:
            print(f"无法解析的样例: {list(unparsed)}")
    
    # 创建时间段
    time_bins = list(range(start_year, end_year + period, period))
    time_labels = [f"{start}-{start+period-1}" for start in time_bins[:-1]]
    
    # 将时间分组
    data['time_period'] = pd.cut(
        data['parsed_time'],
        bins=time_bins,
        labels=time_labels,
        right=False,
        include_lowest=True
    )
    
    # 统计各时间段各类别的数量
    result = pd.crosstab(
        data['time_period'],
        data[category_column],
        margins=True,
        margins_name='总计'
    )
    
    # # 添加时间段内总数
    # result.insert(0, '时段内总数', result.sum(axis=1))
    
    return result


def analyze_time_distribution(
    df: pd.DataFrame,
    time_column: str,
    category_column: str,
    start_year: int = 400,
    end_year: int = 1000,
    period: int = 50,
    top_n: int = 10
) -> Dict:
    """
    分析时间分布的详细信息
    
    返回:
        包含各种统计信息的字典
    """
    # 基础统计
    result = group_by_time_period(
        df, time_column, category_column, 
        start_year, end_year, period, show_details=False
    )
    
    # 解析时间数据
    df_copy = df.copy()
    df_copy['parsed_time'] = df_copy[time_column].apply(parse_time_value)
    valid_data = df_copy[df_copy['parsed_time'].notna()]
    
    analysis = {
        'crosstab': result,
        'time_stats': {
            'min_year': int(valid_data['parsed_time'].min()) if len(valid_data) > 0 else None,
            'max_year': int(valid_data['parsed_time'].max()) if len(valid_data) > 0 else None,
            'mean_year': int(valid_data['parsed_time'].mean()) if len(valid_data) > 0 else None,
            'median_year': int(valid_data['parsed_time'].median()) if len(valid_data) > 0 else None,
        },
        'category_stats': {},
        'top_categories_by_period': {}
    }
    
    # 各类别的统计
    category_counts = valid_data[category_column].value_counts()
    analysis['category_stats'] = {
        'total_categories': len(category_counts),
        'top_categories': category_counts.head(top_n).to_dict(),
        'category_distribution': category_counts.to_dict()
    }
    
    # 每个时期的主要类别
    for period_label in result.index[:-1]:  # 排除"总计"行
        period_data = result.loc[period_label]
        top_in_period = period_data.nlargest(3)
        analysis['top_categories_by_period'][period_label] = top_in_period.to_dict()
    
    return analysis
