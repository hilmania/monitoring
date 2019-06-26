% try
    fileConf = fopen('C:\GConfig\dbconfig.txt');
    serverCon = textscan(fileConf,'%s');
    fclose(fileConf);
    host = serverCon{1}{2};
%     fclose('all');
% catch
%     warning('cant read configuration file!');
% end