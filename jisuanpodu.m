%% 功能：使用远程监控收集的数据计算坡度
%% 将本文件与需要处理的文件放在同一个文件夹下，将从远程监控系统导出的数据打开后另存为.xlsx格式，删除其他的excel文件。
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
path = pwd;
dirOutput = dir(fullfile(path,'*.xlsx'));
fileName = {dirOutput.name};
%% 提示需要处理的文件
fprintf('共找到%d个文件需要处理分别是：',length(fileName));
fprintf('\n');
for i=1:length(fileName)
    fprintf(strcat(fileName{1,i},'\n'));
end
for i=1:length(fileName)
    disp(sprintf('正在处理第%d/%d个文件',i,length(fileName)));  %正在处理第x/X个文件
    filename=fileName{1,i};%%输入数据表格的名称
    f_calculateGradient(filename);
end
system('taskkill /F /IM EXCEL.EXE');% 结束excel进程