%% 使用远程监控收集的数据计算坡度
%% 将本文件与需要处理的文件放在同一个文件夹下，数据为.xls格式且数据为指定格式填写
%% 修改filename的文件名为数据的文件名
%% 处理成功后在当前文件夹下输出podu.xls的文件。

filename='LDYHCS1U0G0008896_20190426143845.xls';%%输入整车配置发动机数据表格
%% 导入驱动电机驱动和发电的所有数据
clc
excelData=xlsread(filename,'sheet1');
[excelRow,excelColumn] = size(excelData);        %%获取数据表中的行列个数
value =  zeros(excelRow,4);                      %建立一个相应行数，1列的矩阵用于存储计算后的数据
invalidDataNum = zeros(1,4);                     %记录数据表前面无效数据的个数。
format short g
%% 仪表车速计算坡度
for row_x = 1: excelRow - 1
    gpsElevationDiffe = excelData(row_x+1,5) - excelData(row_x,5);
    speedSum =  excelData(row_x+1,1) + excelData(row_x,1);
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
    gpsElevationDiffe = excelData(row_x+1,5) - excelData(row_x,5);
    accumulativeMileageDiffe =  excelData(row_x+1,2) - excelData(row_x,2);
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
    gpsElevationDiffe = excelData(row_x+1,5) - excelData(row_x,5);
    gpsSpeedSum =  excelData(row_x+1,3) + excelData(row_x,3);
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
    gpsElevationDiffe = excelData(row_x+1,5) - excelData(row_x,5);
    gpsMileageDiffe =  excelData(row_x+1,4) - excelData(row_x,4);
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
disp('数据处理完毕，请查看当前文件夹下的podu.xls文件');