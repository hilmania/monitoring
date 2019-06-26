function [conn, flag] = OpenConnection()
%# JDBC connector path
% currentfolder = pwd;
% drivername = 'mysql-connector-java-5.0.8-bin.jar';
% javaaddpath([currentfolder '\' drivername]);
%# connection parameteres

% host = '192.168.1.110';
host = 'localhost';
user = 'root';
password = 'disjasad';
dbName = 'disjasad';
jdbcString = sprintf('jdbc:mysql://%s/%s', host, dbName);
jdbcDriver = 'com.mysql.jdbc.Driver';
% logintimeout(jdbcDriver,10);

%# Create the database connection object
conn = ConnectDatabase(dbName, user , password, jdbcDriver, jdbcString);

flag = CheckConnection(conn);
if flag
    display('Connected!')
else
    display('Not Connected!');
end