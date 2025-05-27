# 基础包
import os
import re
import datetime
import argparse

import process_excel as pe
import data_analysis as da

import pdb


if __name__ == "__main__":
    """
    USAGE:
    python main.py -f "../data/山东地区墓志数据.xlsx" -c "书体" --plot --save-plot ../results/plot.html
    """
    args = da.parse_arguments()
    da.analyze_excel_with_args(args)
