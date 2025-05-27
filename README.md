# calligraphAnalysis
## Demand
输入：书法数据统计.excel
输出：多种统计计数结果

## Data Features
表中的列keys包括："朝代	刻石时间	墓志名称	书体	书法风格	横画形态	字形	中宫	起收笔形态	章法	精细程度	界格	地区	官品	规格(长、宽)	墓志规格	官职或出身	墓主身份	图片来源	志主	年代		地区"

## Usage
1、把待处理的 excel文件放到 data目录
2、切换到calligraphAnalysis目录，在命令行运行
```shell
python main.py -f "../data/山东地区墓志数据.xlsx" -c "书体" --plot --save-plot "../results/plot.html"
```
-f 后面，跟"[excel文件]"，-c 后面跟"[待分析列名]"，最后会在命令行打印统计矩阵，并将可视化图表存到--save-plot 指定的目录
