clc
clear
conn = OpenConnection();
ID_Active = RetrieveIdTest(conn)
groupName=RetrieveGroupName(str2double(ID_Active{1}))



% tic
% conn = OpenConnection();
% ID_Active = str2double(RetrieveIdTest(conn));
% fileConf = fopen('C:\GConfig\server.txt');
% serverCon = textscan(fileConf,'%s');
% fclose(fileConf);
% server = ['http://' serverCon{1}{1} '/'];
% params = {'transform','1', 'order','id_test,desc'};
% apiUrl = [server 'apigarjas/api.php/disjas_master_test/'];
% [queryString, ~] = http_paramsToString(params,1);
% readUrl = [apiUrl '?' queryString];
% response = webread(readUrl);
% n=size(response.disjas_master_test,1);
% for k=1:n
%     numID=str2double(response.disjas_master_test(k).id_test);
%     if numID==ID_Active
%         namaKegiatan=response.disjas_master_test(k).nama_kegiatan;
%         break
%     end
% end
% toc