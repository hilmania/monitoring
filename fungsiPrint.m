clc
clear
close all
conn = OpenConnection();
% fileLog = 'C:\ReportAkhir\error.log';
tableName = 'disjas_card_test';
columnName = 'id_user';
id_test = 'id_test';
valueIdTest = RetrieveIdTest(conn);
if strcmp(valueIdTest, 'No Data')
    notif = msgbox('Data belum ada!');
    return
end
rowParticipant ={};

% Query id user to get names
tableUser = 'disjas_master_user';
id = 'id_user';
column = 'nama_user';

full_databases = RetrieveData(conn,tableName,id_test,valueIdTest{1});
databases=full_databases(:,1:end-1);
nData = size(databases,1); %banyaknya peserta
nomorPeserta = {};

for l=1:nData
    valueId = databases{l,3};
    nomorPeserta{l} = ['''' valueId];
    dataParticipant = GetData(conn,tableUser,column,id,valueId);
    rowParticipant{l} = dataParticipant{1}; %variable for nama_user
end

for k=1:nData
    datNames1{k}=[rowParticipant{k}];
end
rowNames=num2cell([1:nData]);

nilaiRerata = cell(1,nData);
temp = [rowNames' nomorPeserta' datNames1' nilaiRerata' databases(:,4:end)];

filename = 'Rekap_Tes_Garjas.xlsx';
A = {'No', 'Nomor Peserta' , 'Nama Peserta', 'Nilai Akhir', 'Lari 3200', 'Pull Up', 'Sit Up', 'Push Up', 'Shuttle Run'};
rekapGarjas = [A;temp];

load tabNilai
nPeserta = size(rekapGarjas, 1);
nColumn = size(rekapGarjas, 2);
nilaiRerata = {};
jumlahA = 0;
jumlahB = 0;

for i=1:nPeserta-1
    [~, usia] = GarjasGetUsia(strrep(rekapGarjas{i+1,2},'''',''));
    rekapGarjas{i+1,3} = [rekapGarjas{i+1,3} ' (' usia ' Thn)'];
    for j=5:nColumn
        switch j
            case 5
                [scoreString, scoreNum] = result2Score('nlongrun', str2double(usia), rekapGarjas{i+1,j}, tabNilai);
                rekapGarjas{i+1,j} = [rekapGarjas{i+1,j} ' (' scoreString ')'];
                jumlahA = jumlahA + scoreNum;
            case 6
                [scoreString, scoreNum] = result2Score('npull_up', str2double(usia), rekapGarjas{i+1,j}, tabNilai);
                rekapGarjas{i+1,j} = [rekapGarjas{i+1,j} ' (' scoreString ')'];
                jumlahB = jumlahB + scoreNum;
            case 7
                [scoreString, scoreNum] = result2Score('nsit_up', str2double(usia), rekapGarjas{i+1,j}, tabNilai);
                rekapGarjas{i+1,j} = [rekapGarjas{i+1,j} ' (' scoreString ')'];
                jumlahB = jumlahB + scoreNum;
            case 8
                [scoreString, scoreNum] = result2Score('npush_up', str2double(usia), rekapGarjas{i+1,j}, tabNilai);
                rekapGarjas{i+1,j} = [rekapGarjas{i+1,j} ' (' scoreString ')'];
                jumlahB = jumlahB + scoreNum;
            case 9
                [scoreString, scoreNum] = result2Score('nshuttle', str2num(usia), rekapGarjas{i+1,j}, tabNilai);
                rekapGarjas{i+1,j} = [rekapGarjas{i+1,j} ' (' scoreString ')'];
                jumlahB = jumlahB + scoreNum;
        end
    end
    nilaiRerata{i} = sprintf('%2.2f',(jumlahA + (jumlahB/4))/2);
    rekapGarjas{i+1,4} = nilaiRerata{i};
    jumlahA = 0;
    jumlahB = 0;
end

try %#ok<TRYNC>
    system('taskkill /F /IM EXCEL.EXE');
end

cfilename = ['C:\ReportAkhir\Report_' date '_' filename];
copyfile(filename, cfilename);
sheet = 1;
xlRange = 'A3';
xlswrite(cfilename, rekapGarjas, sheet, xlRange);

nRow = size(rekapGarjas,1);
nCol = size(rekapGarjas,2);
endIndex = getChara(nCol);
xlRange = ['A3:' endIndex num2str(nRow+2)];
xlsborder(cfilename,'Sheet1',xlRange,'Box',1,2,1,'InsideHorizontal',1,2,1,'InsideVertical',1,2,1);

% currFolder = pwd;
% fullPath = [currFolder '\' filename];
fullPath = cfilename;
hExcel = actxserver('Excel.Application');
hWorkbook = hExcel.workbooks.Open(sprintf('%s', fullPath));
hWorksheet = hWorkbook.Sheets.Item(1);

tgl = date;
fileOutput = ['C:\ReportAkhir\Rekap_Tes_Garjas_' tgl '.pdf'];
hWorksheet.ExportAsFixedFormat('xlTypePDF', fileOutput);
hExcel.ActiveWorkbook.Save;
hExcel.Quit;
delete(hExcel);
open(fileOutput);