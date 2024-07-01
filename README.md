用于解决嘉立创smt贴片bom上传时“上传文件存在位号单元格内容过长”的问题

1. 安装python，并且用pip安装pyopenxlsx
```
pip install pyopenxlsx==3.0.4
```
2. 修改main.py的输入输出文件路径

'in.xlsx' -> 'path/to/new_in.xlsx'

'out.xlsx' -> 'path/to/new_out.xlsx'

3. 执行
```
python main.py
```

PS:请使用默认配置导出的xlsx
