# 基础包
from pathlib import Path

# 数据处理包
import pandas as pd
import numpy as np

# Excel文件处理包
import openpyxl  # 处理.xlsx文件
import xlrd      # 读取.xls文件（旧版Excel）
import xlwt      # 写入.xls文件（旧版Excel）
import xlsxwriter # 另一个写入.xlsx的库，功能更强大

# 其他实用包
import logging
from typing import List, Dict, Optional, Union

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

# Parse excel
def read_excel_advanced(
    filename: str,
    sheet_name: Optional[Union[str, int, List]] = None,
    header: Optional[Union[int, List[int]]] = 0,
    usecols: Optional[Union[str, List[str], List[int]]] = None,
    dtype: Optional[Dict[str, type]] = None,
    skiprows: Optional[Union[int, List[int]]] = None,
    nrows: Optional[int] = None,
    parse_dates: Optional[Union[bool, List[str], Dict]] = True,
    date_parser: Optional[callable] = None,
    na_values: Optional[Union[str, List[str], Dict]] = None,
    converters: Optional[Dict[str, callable]] = None
) -> Union[pd.DataFrame, Dict[str, pd.DataFrame]]:
    """
    高级Excel读取函数，支持更多参数配置
    
    参数:
        filename: Excel文件路径
        sheet_name: 工作表名称或索引
        header: 列名所在行
        usecols: 要读取的列
        dtype: 列的数据类型
        skiprows: 跳过的行数
        nrows: 读取的行数
        parse_dates: 日期解析配置
        date_parser: 自定义日期解析函数
        na_values: 空值定义
        converters: 列转换函数
    
    返回:
        DataFrame或字典
    """
    try:
        # 构建读取参数
        read_params = {
            'io': filename,
            'sheet_name': sheet_name,
            'header': header,
            'engine': 'openpyxl'
        }
        
        # 添加可选参数
        if usecols is not None:
            read_params['usecols'] = usecols
        if dtype is not None:
            read_params['dtype'] = dtype
        if skiprows is not None:
            read_params['skiprows'] = skiprows
        if nrows is not None:
            read_params['nrows'] = nrows
        if parse_dates is not None:
            read_params['parse_dates'] = parse_dates
        if date_parser is not None:
            read_params['date_parser'] = date_parser
        if na_values is not None:
            read_params['na_values'] = na_values
        if converters is not None:
            read_params['converters'] = converters
        
        # 读取数据
        result = pd.read_excel(**read_params)
        
        # 如果返回的是字典，对每个DataFrame进行清洗
        if isinstance(result, dict):
            for sheet, df in result.items():
                result[sheet] = clean_dataframe(df)
        else:
            result = clean_dataframe(result)
        
        return result
        
    except Exception as e:
        logger.error(f"高级读取失败: {str(e)}")
        raise


# 批量处理Excel文件
def batch_read_excel_files(
    folder_path: str,
    pattern: str = "*.xlsx",
    combine: bool = False
) -> Union[Dict[str, pd.DataFrame], pd.DataFrame]:
    """
    批量读取文件夹中的Excel文件
    
    参数:
        folder_path: 文件夹路径
        pattern: 文件匹配模式
        combine: 是否合并所有数据
    
    返回:
        字典或合并后的DataFrame
    """
    folder = Path(folder_path)
    excel_files = list(folder.glob(pattern))
    
    if not excel_files:
        raise ValueError(f"在 {folder_path} 中没有找到匹配 {pattern} 的文件")
    
    print(f"找到 {len(excel_files)} 个Excel文件")
    
    all_data = {}
    
    for file in excel_files:
        print(f"读取: {file.name}")
        try:
            df = read_excel_advanced(str(file), sheet_name=0)
            all_data[file.stem] = df
        except Exception as e:
            logger.error(f"读取 {file.name} 失败: {str(e)}")
    
    if combine:
        # 合并所有DataFrame
        combined_df = pd.concat(all_data.values(), ignore_index=True, sort=False)
        print(f"合并后数据: {combined_df.shape[0]} 行 × {combined_df.shape[1]} 列")
        return combined_df
    else:
        return all_data
    

def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    清洗DataFrame数据
    
    参数:
        df: 原始DataFrame
    
    返回:
        清洗后的DataFrame
    """
    # 删除全空行
    df = df.dropna(how='all')
    
    # 删除全空列
    df = df.dropna(axis=1, how='all')
    
    # 去除列名的前后空格
    df.columns = df.columns.str.strip()
    
    # 去除字符串类型数据的前后空格
    str_columns = df.select_dtypes(include=['object']).columns
    for col in str_columns:
        df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
    
    return df

