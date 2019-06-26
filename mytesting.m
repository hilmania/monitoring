clc
clear all


ipAddr = '192.168.1.20';
user = 'root';
password = 'MetaVision916110';
dbName = 'garjasad';
conf={ipAddr;user;password;dbName};

namaFile = 'myserver.txt';

% if exist(namaFile, 'file') == 2
%     display('File configuration already exist!')
% else
%     try
        fid=fopen(namaFile, 'w');
        fprintf(fid, ['%s\r\n%s\r\n%s\r\n%s'],conf{1},conf{2},conf{3},conf{4});
        fclose(fid);
%     catch
%         warning('Cannot write a file!');
%     end
% end
fclose('all');


% fprintf( ['The results of test %d are such that %d of the ', ...
% 'cats are older than %d years old.\nThe results of test ', ... 
% '%d are such that %d of the dogs are older than %d years ', ...
% 'old.\nThe results of test %d are such that %d of the ', ...
% 'fish are older than %d years old.\n'], t1, cats, age, ...
% t2, dogs, age, t3, fish, age);