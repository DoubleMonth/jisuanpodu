功能：使用远程监控收集的数据计算坡度

2019-05-16之前的程序请使用LDYHCS1U0G0008896_20190426143845.xls作为模板。

2019-05-16版本程序请使用LDYECS744J0008255_20190516142054.xlsx作为模板进行。

2019-05-16 将程序修改为直接使用远程监控导出的数据（需要进行一下格式转换，请参阅文件中说明），生成的podu文件中有计算后的数据，也拷贝出计算需要的数据以备手动计算使用。

2019-05-17 在生成的EXCEL表中添加一列时间

2019-05-18 添加生成曲线图功能。在生成的EXCEL表中添加序号，EXCEL生成的数据保存在"EXCEL表名_output"文件夹啊，生成的EXCEL表和图片也使用相应的文件名进行命名。
2019-06-09 可处理当前文件夹下的所有合法Excel文件。并分别输出在对应的文件夹中。在EXCEL中插入折线图。
2019-06-20 修改提示文件已存在，需要替换以前文件的提示；修改出现warning xlswrite 警告。