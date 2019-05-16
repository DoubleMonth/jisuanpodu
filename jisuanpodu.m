%% ���ܣ�ʹ��Զ�̼���ռ������ݼ����¶�
%% �����ļ�����Ҫ������ļ�����ͬһ���ļ����£�����Զ�̼��ϵͳ���������ݴ򿪺����Ϊ.xlsx��ʽ
%% --��Ϊ����������Ϊ�ı���ʽ��excel�ļ���MATLAB����ʱ�����
%% �޸�filename���ļ���Ϊ���ݵ��ļ���
%% �������Ҫ�������ݣ���ע�͵��˲��㷨���֣������Ҫ���˵����ݲ���0.2�������˲��㷨���޸���Ӧ����
%% ����ɹ����ڵ�ǰ�ļ��������podu.xls���ļ���
%% ע�⣺���벻Ҫ�����ļ����������ĵ�·����ʹ�ã��������ִ��󣻳�������ǰ��ر��Ѿ��򿪵�podu.xls�ļ���

%% ��   ����(20190516)V1.0
%% ��   �ߣ�
%% �޸�ʱ�䣺2019-5-16
clear           %% ��������ռ��е����б�����
clc
disp('���ڴ������ݣ����Ժ�........');
filename='LDYECS744J0008255_20190516142054.xlsx';%%�������ݱ�������
[excelData,str] = xlsread(filename,1);               %��ȡԭʼ���ݱ��е����ݣ�strΪ���ݱ��е��ַ���dataΪ���ݱ��е�����
[excelRow,excelColumn] = size(excelData);        %%��ȡ���ݱ��е����и���
value =  zeros(excelRow,4);                      %����һ����Ӧ������1�еľ������ڴ洢����������
invalidDataNum = zeros(1,4);                     %��¼���ݱ�ǰ����Ч���ݵĸ�����
[m,n] = size(str);                              %% ���ݱ����ַ��ĸ���
needStr = {'����','�ۼ����','GPS����','GPS���','GPS����'}; %% �����¶���Ҫ��������
needStrStationIn_value = zeros(1,5);                        %% ����������ԭʼ���ݱ��е�λ��

%% �ҳ���Ҫ����������ԭʼ���ݱ��е�λ��
for i = 1 :n                        
    for j = 1: 5
        if strcmp(str(1,i),needStr(1,j))>0
            needStrStationIn_value(1,j) = i-1;      %% -1����Ϊ��ԭ���ݱ��е�һ��Ϊʱ�䣬MATLAB��ȡ�����ݾ�����û����һ�С�excelData��������������ԭʼexcel��һ�У���һ�С�
        end
    end
end
format short g                                      %%������ʾ��ʽ
%% �Ǳ��ټ����¶�
for row_x = 1: excelRow - 1
    gpsElevationDiffe = excelData(row_x+1,needStrStationIn_value(1,5)) - excelData(row_x,needStrStationIn_value(1,5));
    speedSum =  excelData(row_x+1,needStrStationIn_value(1,1)) + excelData(row_x,needStrStationIn_value(1,1));
    if speedSum == 0                                                        % speedSum=0ʱ����ĸΪ0����Ч����
        if invalidDataNum(1,1) == 0                                         % ��û�м�¼��Ч���ݸ���ʱ��Ч���ݵ�λ�����0
            value(row_x,1) = 0;
        else
            value(row_x,1) = value(row_x-1,1);                          %��¼��Ч���ݸ���������һ�����ݽ������
        end
    else                                                                   %��Ч����
        if invalidDataNum(1,1) == 0                                        
            invalidDataNum(1,1) = row_x;                                  %��û�м�¼��Ч����ʱ��¼����Ч���ݵĸ���
        end     
        mid_value_2 = asind(gpsElevationDiffe/(speedSum/2*5/3600*1000)); 
        if isreal(mid_value_2)                                              
            podu =  tand(mid_value_2);
            value(row_x,1) = podu;                                       %д�������
        else
            value(row_x,1) = value(row_x-1,1);                         %%���ָ���ʱ��Ϊ���ݳ���ʹ����һ���������
        end
    end
end
%% �ۼ���̼����¶�
for row_x = 1: excelRow - 1
    gpsElevationDiffe = excelData(row_x+1,needStrStationIn_value(1,5)) - excelData(row_x,needStrStationIn_value(1,5));
    accumulativeMileageDiffe =  excelData(row_x+1,needStrStationIn_value(1,2)) - excelData(row_x,needStrStationIn_value(1,2));
    if accumulativeMileageDiffe == 0                                        %%�ۼ���̲����0���������ʱʹ����һ�е����ݽ������
        if invalidDataNum(1,2) == 0                                         % ��û�м�¼��Ч���ݸ���ʱ��Ч���ݵ�λ�����0
            value(row_x,2) = 0;
        else
            value(row_x,2) = value(row_x-1,2);
        end
    else   
        if invalidDataNum(1,2) == 0                                         % ��û�м�¼��Ч���ݸ���ʱ��Ч���ݵ�λ�����0
            invalidDataNum(1,2) = row_x;
        end
        mid_value_2 = asind(gpsElevationDiffe/accumulativeMileageDiffe/1000); 
            podu =  tand(mid_value_2);
            value(row_x,2) = podu;                                         %д�������
    end
end
%% GPS���ټ����¶�
for row_x = 1: excelRow - 1
    gpsElevationDiffe = excelData(row_x+1,needStrStationIn_value(1,5)) - excelData(row_x,needStrStationIn_value(1,5));
    gpsSpeedSum =  excelData(row_x+1,needStrStationIn_value(1,3)) + excelData(row_x,needStrStationIn_value(1,3));
    if gpsSpeedSum == 0                                                    %%�ٶȻ����0���������ʱʹ����һ�е����ݽ������
        if invalidDataNum(1,3) == 0                                         % ��û�м�¼��Ч���ݸ���ʱ��Ч���ݵ�λ�����0
            value(row_x,3) = 0;
        else
            value(row_x,3) = value(row_x-1,3);
        end
    else   
        if invalidDataNum(1,3) == 0                                         % ��û�м�¼��Ч���ݸ���ʱ��Ч���ݵ�λ�����0
            invalidDataNum(1,3) = row_x;
        end
        mid_value_2 = asind(gpsElevationDiffe/(gpsSpeedSum/2*5/3600*1000)); 
        if isreal(mid_value_2)                                              %%���ָ���ʱ��Ϊ���ݳ���ʹ����һ���������
            podu =  tand(mid_value_2);
            value(row_x,3) = podu;                                        %д�������
        else
            value(row_x,3) = value(row_x-1,3);
        end
    end
end
%% GPS��̼����¶�
for row_x = 1: excelRow - 1
    gpsElevationDiffe = excelData(row_x+1,needStrStationIn_value(1,5)) - excelData(row_x,needStrStationIn_value(1,5));
    gpsMileageDiffe =  excelData(row_x+1,needStrStationIn_value(1,4)) - excelData(row_x,needStrStationIn_value(1,4));
    if gpsMileageDiffe == 0                                                %%�ٶȻ����0���������ʱʹ����һ�е����ݽ������
        if invalidDataNum(1,4) == 0                                        % ��û�м�¼��Ч���ݸ���ʱ��Ч���ݵ�λ�����0
            value(row_x,4) = 0;
        else
            value(row_x,4) = value(row_x-1,4);
        end
    else   
        if invalidDataNum(1,4) == 0                                        % ��û�м�¼��Ч���ݸ���ʱ��Ч���ݵ�λ�����0
            invalidDataNum(1,4) = row_x;
        end
        mid_value_2 = gpsElevationDiffe/gpsMileageDiffe/1000; 
            podu =  tan(mid_value_2);
            value(row_x,4) = podu;                                        %д�������
    end
end
%% �˲��㷨--ȥ��>0.2��<-0.2�����ݣ�����һ���������
for i = 1:4
    for j = 2:excelRow
        if value(j,i)>0.2||value(j,i)<-0.2
            value(j,i) = value(j-1,i);
        end
    end
end
%% �����ǰ����podu�ļ�����ɾ��
if exist('podu.xls')   
    delete('podu.xls');
end
value_2 = value(max(invalidDataNum(:)):excelRow,1:4);                               %%ȡ����������Ч���ݣ�������Ч����
colname={'�Ǳ��ټ����¶�','�ۼ���̼����¶�','GPS���ټ����¶�','GPS��̼����¶�'};    %%����ÿһ�е���������
xlswrite('podu.xls', colname, 'sheet1','A1');
xlswrite('podu.xls',value_2, 'sheet1','A2');
%% ��������Ҫ�����ݿ���һ�ݷ���Sheet2�У��Ա��ֶ�����ʱʹ�á�
xlswrite('podu.xls',str(:,1), 'sheet2','A1');
xlswrite('podu.xls',needStr, 'sheet2','B1');
xlswrite('podu.xls',excelData(:,needStrStationIn_value(1,1)), 'sheet2','B2');
xlswrite('podu.xls',excelData(:,needStrStationIn_value(1,2)), 'sheet2','C2');
xlswrite('podu.xls',excelData(:,needStrStationIn_value(1,3)), 'sheet2','D2');
xlswrite('podu.xls',excelData(:,needStrStationIn_value(1,4)), 'sheet2','E2');
xlswrite('podu.xls',excelData(:,needStrStationIn_value(1,5)), 'sheet2','F2');
%% Sheet������
path = pwd;
filePath = fullfile(path,'podu.xls');
e = actxserver('Excel.Application');
ewb = e.workbooks.Open(filePath);
ewb.Worksheets.Item(1).Name = '������¶�';
ewb.Worksheets.Item(2).Name = '�����¶�ʹ�õ�����';
ewb.Save 
ewb.Close(false)
e.Quit
%% ���ݴ�����ϣ������ʾ��Ϣ
disp('���ݴ�����ϣ���鿴��ǰ�ļ����µ�podu.xls�ļ�');
filePath %% �ļ�����λ��