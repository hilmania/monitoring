function [conn, flag] = OpenConnection()

try
    fileConf = fopen('C:\GConfig\server.txt');
    serverCon = textscan(fileConf,'%s');
    fclose(fileConf);
    host = serverCon{1}{1};
catch
    host='192.168.1.11';
end
user = 'root';
password = 'MetaVision916110';
dbName = 'garjasad';
jdbcString = sprintf('jdbc:mysql://%s/%s', host, dbName);
jdbcDriver = 'com.mysql.jdbc.Driver';
% logintimeout(jdbcDriver,10);

%# Create the database connection object
conn = ConnectDatabase(dbName, user , password, jdbcDriver, jdbcString);

flag = CheckConnection(conn);
% if flag
%     display('Connected!')
% else
%     display('Not Connected!');
% end