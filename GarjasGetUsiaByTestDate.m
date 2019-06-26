% function to get usia from barcode/id_user provided
% @param id_user (barcode)

function [flag, usia] = GarjasGetUsiaByTestDate(id_user)
try
    fileConf = fopen('C:\GConfig\server.txt');
    serverCon = textscan(fileConf,'%s');
    fclose(fileConf);
    server = ['http://' serverCon{1}{1} '/'];
    apiUrl = [server 'apigarjas/api.php/disjas_master_user/'];
    readUrl = [apiUrl id_user];
    data = webread(readUrl);
    usia = data.usia;
    flag = true;
catch
    warning('Connection refused!');
    flag = false;
    usia = '0';
end