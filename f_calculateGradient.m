function  f_calculateGradient( r_fileName )
%UNTITLED Summary of this function goes here
%   Detailed explanation goes here
samplingTime = 5; %%采样时间
path = pwd;
%% 输出提示信息
disp(strcat('正在处理',r_fileName));
disp('请稍候......');
%% 处理Excel表格
[excelData,str] = xlsread(r_fileName,1);               %读取原始数据表中的数据：str为数据表中的字符，data为数据表中的数据
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
        mid_value_2 = asind(gpsElevationDiffe/(speedSum/2*samplingTime/3600*1000)); 
        if isreal(mid_value_2)                                              
            podu = tand(mid_value_2);
            value(row_x,1) = podu *100;                                       %写入矩阵中
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
            value(row_x,2) = podu*100;                                         %写入矩阵中
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
        mid_value_2 = asind(gpsElevationDiffe/(gpsSpeedSum/2*samplingTime/3600*1000)); %% 注意修改采样时间
        if isreal(mid_value_2)                                              %%出现复数时认为数据出错，使用上一个数据填充
            podu =  tand(mid_value_2);
            value(row_x,3) = podu*100;                                        %写入矩阵中
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
        mid_value_2 = asind(gpsElevationDiffe/gpsMileageDiffe/1000); 
            podu = tand(mid_value_2);
            value(row_x,4) = podu*100;                                        %写入矩阵中
    end
end
%% 滤波算法--去除>0.2和<-0.2的数据，用上一个数据填充
for i = 1:4
    for j = 2:excelRow
        if value(j,i)>20||value(j,i)<-20
            value(j,i) = value(j-1,i);
        end
    end
end
i = find('.'==r_fileName);
imname = r_fileName(1:i-1); %% imname为不带后缀文件名称 
outFile = strcat(imname,'_output');
if exist(outFile)   %% 如果存在output文件夹，先删除
     rmdir (outFile,'s');
end
mkdir(outFile);%% 创建一个Output文件夹
cd(fullfile(path,outFile));       %%进入output目录
poduFile = strcat(imname,'_podu.xls'); %%组成带excle文件名的podu文件名
value_2 = value(max(invalidDataNum(:)):excelRow,1:4);                               %%取出矩阵中有效数据，丢弃无效数据
colname={'序号','时间','仪表车速计算坡度','累计里程计算坡度','GPS车速计算坡度','GPS里程计算坡度'};    %%增加每一列的数据名称
warning off MATLAB:xlswrite:AddSheet;   %%防止出现warning警告 
xlswrite(poduFile, colname, 'sheet1','A1');
xuhao = linspace(1,m-max(invalidDataNum(:)),m-max(invalidDataNum(:)));
xlswrite(poduFile, xuhao', 'sheet1','A2');                %%序号
xlswrite(poduFile,str(max(invalidDataNum(:))+1:m,1), 'sheet1','B2');              %%时间
xlswrite(poduFile,value_2, 'sheet1','C2');                    %%计算后的数据
%% 将计算需要的数据拷贝一份放入Sheet2中，以备手动计算时使用。
xlswrite(poduFile,str(:,1), 'sheet2','A1');
xlswrite(poduFile,needStr, 'sheet2','B1');
xlswrite(poduFile,excelData(:,needStrStationIn_value(1,1)), 'sheet2','B2');
xlswrite(poduFile,excelData(:,needStrStationIn_value(1,2)), 'sheet2','C2');
xlswrite(poduFile,excelData(:,needStrStationIn_value(1,3)), 'sheet2','D2');
xlswrite(poduFile,excelData(:,needStrStationIn_value(1,4)), 'sheet2','E2');
xlswrite(poduFile,excelData(:,needStrStationIn_value(1,5)), 'sheet2','F2');
colname2={'=TAN(ASIN((F3-F2)/((B2+B3)/2*10/3600*1000)))','=TAN(ASIN((F3-F2)/(C3-C2)/1000))','=TAN(ASIN((F3-F2)/((D3+D2)/2*10/3600*1000)))','=TAN(ASIN((F3-F2)/(E3-E2)/1000))'};
xlswrite(poduFile,colname2, 'sheet2','G2');

%% Sheet重命名
path = pwd;
filePath = fullfile(path,poduFile);
e = actxserver('Excel.Application');
ewb = e.workbooks.Open(filePath);
ewb.Worksheets.Item(1).Name = '计算的坡度';
ewb.Worksheets.Item(2).Name = '计算坡度使用的数据';
ewb.Save 
ewb.Close(false)
e.Quit

%% 绘制折线图
fh=figure;
set(fh,'visible','off');    %% 绘制折线时不显示figure窗口
plot(xuhao',value_2(:,1),'y-');
hold on;
plot(xuhao',value_2(:,2),'b-');
hold on;
plot(xuhao',value_2(:,3),'g-');
hold on;
plot(xuhao',value_2(:,4),'r-');
hold on;
grid on;%%显示网格线
legend('仪表车速计算坡度','累计里程计算坡度','GPS车速计算坡度','GPS里程计算坡度');
titleFile = strcat(imname,'计算坡度'); %%组成带excle文件名的podu文件名
title(titleFile);
xlabel('时间顺序');
ylabel('坡度');

%% 将折线图插入EXCE中
print -dbitmap
spwd = [pwd '\'];
filespec_user=[spwd poduFile];
try
Excel=actxGetRunningServer('Excel.Application');
catch
Excel = actxserver('Excel.Application');
end;
%设置Excel属性为可见
set(Excel, 'Visible', 0);
Workbooks = Excel.Workbooks;
%若测试文件存在，打开该测试文件，否则，新建一个工作簿，并保存，文件名为测试.Excel
if exist(filespec_user,'file');
Workbook = invoke(Workbooks,'Open',filespec_user);
else
Workbook = invoke(Workbooks, 'Add');
Workbook.SaveAs(filespec_user);
end
%返回工作表句柄
Sheets = Excel.ActiveWorkBook.Sheets;
%返回第一个表格句柄
sheet1 = get(Sheets, 'Item', 1);
%激活第一个表格
invoke(sheet1, 'Activate');
%如果当前工作表中有图形存在，通过循环将图形全部删除
Shapes=Excel.ActiveSheet.Shapes;
if Shapes.Count~=0;
for i=1:Shapes.Count;
Shapes.Item(1).Delete;
end;
end;

%将图形复制到粘贴板
hgexport(fh, '-clipboard');
%将图形粘贴到当前表格的A5:B5栏里
Excel.ActiveSheet.Range('H5:I5').Select;
Excel.ActiveSheet.Paste;
%删除图形句柄
% 
Workbook.Save();%% 保存文件
% Workbook.SaveAs(filespec_user);%% 保存文件
delete(Workbook);

%% 保存生成的折线图  先删除以前存在的图片
pngFile = strcat(imname,'_tupian.png'); %%组成带excle文件名的podu文件名
figFile = strcat(imname,'_tupian.fig'); %%组成带excle文件名的podu文件名
% if exist('tupian.png')   
%     delete('tupian.png');
% end
% if exist('tupian.fig')   
%     delete('tupian.fig');
% end
saveas(gcf,pngFile);
saveas(gcf,figFile);
delete(fh);
%% 数据处理完毕，输出提示信息
disp('数据处理完毕，请查看当前文件夹下的');
disp(poduFile);
filePath    %% 文件所在位置
cd ..       %%退出output目录

end

