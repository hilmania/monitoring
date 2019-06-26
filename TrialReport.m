function TrialReport
clc
clear all
close all
conn = OpenConnection();

tableName = 'disjas_card_test';
columnName = 'id_user';
id_test = 'id_test';
valueIdTest = RetrieveIdTest(conn);
rowParticipant ={};

% Query id user to get names
tableUser = 'disjas_master_user';
id = 'id_user';
column = 'nama_user';

databases = RetrieveData(conn,tableName,id_test,valueIdTest{1})
nData = size(databases,1);
nomorPeserta = {};

for l=1:nData
    valueId = databases{l,3};
    nomorPeserta{l} = ['''' valueId];
    dataParticipant = GetData(conn,tableUser,column,id,valueId);
    rowParticipant{l} = dataParticipant{1} %variable for nama_user
end

for k=1:nData
    datNames1{k}=[rowParticipant{k}];
end
rowNames=num2cell([1:nData]);

temp = [rowNames' nomorPeserta' datNames1' databases(1:end,4:end-2)]

filename = 'Rekap_Tes_Garjas.xlsx';
A = {'No', 'Nomor Peserta' , 'Nama Peserta', 'Pull Up', 'Push Up', 'Sit Up', 'Long Run', 'Shuttle Run'};
rekapGarjas = [A;temp];

xlswrite(filename, rekapGarjas);

% nFile = xlsread(filename);
nRow = size(rekapGarjas,1);
nCol = size(rekapGarjas,2);
endIndex = getChara(nCol);
xlRange = ['A1:' endIndex num2str(nRow+1)];
xlsborder(filename,'Sheet1',xlRange,'Box',1,2,1,'InsideHorizontal',1,2,1,'InsideVertical',1,2,1);

currFolder = pwd;
fullPath = [currFolder '\' filename];
hExcel = actxserver('Excel.Application');
hWorkbook = hExcel.workbooks.Open(sprintf('%s', fullPath));
hWorksheet = hWorkbook.Sheets.Item(1);

fileOutput = [currFolder '\Rekap_Tes_Garjas.pdf'];
hWorksheet.ExportAsFixedFormat('xlTypePDF', fileOutput);