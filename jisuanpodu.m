%% 功能：使用远程监控收集的数据计算坡度
%% 将本文件与需要处理的文件放在同一个文件夹下，将从远程监控系统导出的数据打开后另存为.xlsx格式
%% --因为导出的数据为文本格式的excel文件，MATLAB处理时会出错。
%% 修改filename的文件名为数据的文件名
%% 如果不需要过滤数据，请注释掉滤波算法部分，如果需要过滤的数据不是0.2，请在滤波算法中修改相应数据
%% 处理成功后在当前文件夹下输出podu.xls的文件。
%% 注意：：请不要将本文件放在有中文的路径下使用，否则会出现错误；程序运行前请关闭已经打开的podu.xls文件。

%% 版   本：(20190516)V1.0
%% 作   者：
%% 修改时间：2019-5-16
clear           %% 清除工作空间中的所有变量。
clc
disp('正在处理数据，请稍候........');
filename='LDYECS744J0008255_20190516142054.xlsx';%%输入数据表格的名称
[excelData,str] = xlsread(filename,1);               %读取原始数据表中的数据：str为数据表中的字符，data为数据表中的数据
[excelRow,excelColumn] = size(excelData);        %%获取数据表中的行列个数
value =  zeros(excelRow,4);                      %建立一个相应行数，1列的矩阵用于存储计算后的数据
invalidDataNum = zeros(1,4);                     %记录数据表前面无效数据的个数。
[m,n] = size(str);                              %% 数据表中字符的个数
needStr = {'车速','累计里程','GPS车速','GPS里程','GPS海拔'}; %% 计算坡度需要的数据项
needStrStationIn_value = zeros(1,5);                        %% 各数据项在原始数据表中的位置

%% 找出需要的数据项在原始数据表中的位置
for i = 1 :n                        
    for j = 1: 5
        if strcmp(str(1,i),needStr(1,j))>0
            needStrStationIn_value(1,j) = i-1;      %% -1是因为在原数据表中第一列为时间，MATLAB读取后数据矩阵中没有这一列。excelData的行数和列数比原始excel少一行，少一列。
        end
    end
end
format short g                                      %%设置显示格式
%% 仪表车速计算坡度
for row_x = 1: excelRow - 1
    gpsElevationDiffe = excelData(row_x+1,needStrStationIn_value(1,5)) - excelData(row_x,needStrStationIn_value(1,5));
    speedSum =  excelData(row_x+1,needStrStationIn_value(1,1)) + excelData(row_x,needStrStationIn_value(1,1));
    if speedSum == 0                                                        % speedSum=0时，分母为0，无效数据
        if invalidDataNum(1,1) == 0                                         % 还没有记录无效数据个数时无效数据的位置填充0
            value(row_x,1) = 0;
        else
            value(row_x,1) = value(row_x-1,1);                          %记录无效数据个数后用上一个数据进行填充
        end
    else                                                                   %有效数据
        if invalidDataNum(1,1) == 0                                        
            invalidDataNum(1,1) = row_x;                                  %还没有记录无效数据时记录下无效数据的个数
        end     
        mid_value_2 = asind(gpsElevationDiffe/(speedSum/2*5/3600*1000)); 
        if isreal(mid_value_2)                                              
            podu =  tand(mid_value_2);
            value(row_x,1) = podu;                                       %写入矩阵中
        else
            value(row_x,1) = value(row_x-1,1);                         %%出现复数时认为数据出错，使用上一个数据填充
        end
    end
end
%% 累计里程计算坡度
for row_x = 1: excelRow - 1
    gpsElevationDiffe = excelData(row_x+1,needStrStationIn_value(1,5)) - excelData(row_x,needStrStationIn_value(1,5));
    accumulativeMileageDiffe =  excelData(row_x+1,needStrStationIn_value(1,2)) - excelData(row_x,needStrStationIn_value(1,2));
    if accumulativeMileageDiffe == 0                                        %%累计里程差等于0的情况，此时使用上一行的数据进行填充
        if invalidDataNum(1,2) == 0                                         % 还没有记录无效数据个数时无效数据的位置填充0
            value(row_x,2) = 0;
        else
            value(row_x,2) = value(row_x-1,2);
        end
    else   
        if invalidDataNum(1,2) == 0                                         % 还没有记录无效数据个数时无效数据的位置填充0
            invalidDataNum(1,2) = row_x;
        end
        mid_value_2 = asind(gpsElevationDiffe/accumulativeMileageDiffe/1000); 
            podu =  tand(mid_value_2);
            value(row_x,2) = podu;                                         %写入矩阵中
    end
end
%% GPS车速计算坡度
for row_x = 1: excelRow - 1
    gpsElevationDiffe = excelData(row_x+1,needStrStationIn_value(1,5)) - excelData(row_x,needStrStationIn_value(1,5));
    gpsSpeedSum =  excelData(row_x+1,needStrStationIn_value(1,3)) + excelData(row_x,needStrStationIn_value(1,3));
    if gpsSpeedSum == 0                                                    %%速度会出现0的情况，此时使用上一行的数据进行填充
        if invalidDataNum(1,3) == 0                                         % 还没有记录无效数据个数时无效数据的位置填充0
            value(row_x,3) = 0;
        else
            value(row_x,3) = value(row_x-1,3);
        end
    else   
        if invalidDataNum(1,3) == 0                                         % 还没有记录无效数据个数时无效数据的位置填充0
            invalidDataNum(1,3) = row_x;
        end
        mid_value_2 = asind(gpsElevationDiffe/(gpsSpeedSum/2*5/3600*1000)); 
        if isreal(mid_value_2)                                              %%出现复数时认为数据出错，使用上一个数据填充
            podu =  tand(mid_value_2);
            value(row_x,3) = podu;                                        %写入矩阵中
        else
            value(row_x,3) = value(row_x-1,3);
        end
    end
end
%% GPS里程计算坡度
for row_x = 1: excelRow - 1
    gpsElevationDiffe = excelData(row_x+1,needStrStationIn_value(1,5)) - excelData(row_x,needStrStationIn_value(1,5));
    gpsMileageDiffe =  excelData(row_x+1,needStrStationIn_value(1,4)) - excelData(row_x,needStrStationIn_value(1,4));
    if gpsMileageDiffe == 0                                                %%速度会出现0的情况，此时使用上一行的数据进行填充
        if invalidDataNum(1,4) == 0                                        % 还没有记录无效数据个数时无效数据的位置填充0
            value(row_x,4) = 0;
        else
            value(row_x,4) = value(row_x-1,4);
        end
    else   
        if invalidDataNum(1,4) == 0                                        % 还没有记录无效数据个数时无效数据的位置填充0
            invalidDataNum(1,4) = row_x;
        end
        mid_value_2 = gpsElevationDiffe/gpsMileageDiffe/1000; 
            podu =  tan(mid_value_2);
            value(row_x,4) = podu;                                        %写入矩阵中
    end
end
%% 滤波算法--去除>0.2和<-0.2的数据，用上一个数据填充
for i = 1:4
    for j = 2:excelRow
        if value(j,i)>0.2||value(j,i)<-0.2
            value(j,i) = value(j-1,i);
        end
    end
end
%% 如果以前存在podu文件，先删除
if exist('podu.xls')   
    delete('podu.xls');
end
value_2 = value(max(invalidDataNum(:)):excelRow,1:4);                               %%取出矩阵中有效数据，丢弃无效数据
colname={'仪表车速计算坡度','累计里程计算坡度','GPS车速计算坡度','GPS里程计算坡度'};    %%增加每一列的数据名称
xlswrite('podu.xls', colname, 'sheet1','A1');
xlswrite('podu.xls',value_2, 'sheet1','A2');
%% 将计算需要的数据拷贝一份放入Sheet2中，以备手动计算时使用。
xlswrite('podu.xls',str(:,1), 'sheet2','A1');
xlswrite('podu.xls',needStr, 'sheet2','B1');
xlswrite('podu.xls',excelData(:,needStrStationIn_value(1,1)), 'sheet2','B2');
xlswrite('podu.xls',excelData(:,needStrStationIn_value(1,2)), 'sheet2','C2');
xlswrite('podu.xls',excelData(:,needStrStationIn_value(1,3)), 'sheet2','D2');
xlswrite('podu.xls',excelData(:,needStrStationIn_value(1,4)), 'sheet2','E2');
xlswrite('podu.xls',excelData(:,needStrStationIn_value(1,5)), 'sheet2','F2');
%% Sheet重命令
path = pwd;
filePath = fullfile(path,'podu.xls');
e = actxserver('Excel.Application');
ewb = e.workbooks.Open(filePath);
ewb.Worksheets.Item(1).Name = '计算的坡度';
ewb.Worksheets.Item(2).Name = '计算坡度使用的数据';
ewb.Save 
ewb.Close(false)
e.Quit
%% 数据处理完毕，输出提示信息
disp('数据处理完毕，请查看当前文件夹下的podu.xls文件');
filePath %% 文件所在位置