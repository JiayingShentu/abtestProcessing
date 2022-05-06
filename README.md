##爱奇艺弹幕产品AB测试数据处理流程     
###常规处理操作      
1. 将AB实验数据结果的&nbsp;.xlsx&nbsp;文件放入文件夹中,重命名为&nbsp; origin.xlsx&nbsp;;
2. 确保文件夹中无&nbsp; result.xlsx&nbsp;文件;
3. 运行&nbsp; processing.py &nbsp;文件即可.
###前期配置     
1. 下载python&nbsp; <https://www.python.org/>;
2. 拥有一个python开发工具，例如Python自带IDLE，pycharm，vscode等;
3. 下载&nbsp; pandas库&nbsp; 和&nbsp;openpyxl库&nbsp; `pip install pandas`&nbsp;&&nbsp;`pip install openpyxl`
###个性化修改      
1. processing.py &nbsp; line7 setIndex的两个参数应该修改为origin.xlsx表中的表示时间和AB实验的两个字段名称;    
例：`setIndex=['日期','AB实验']`
2. 不建议修改指标，如果非要新增指标，需要相对应地新增 function.py 中的 &nbsp;type1&nbsp;和&nbsp;type2&nbsp; 数组元素，已有指标包括 function.py 中&nbsp;title数组&nbsp;中的所有字段

