function [namaFile, flag] = GarjasSaveOffline(id_user, nilai, station, namaTestor)

header1 = 'IDUser';
header2 = 'NamaUser';
header3 = 'Nilai';
tanggal = date;
namaFile = [tanggal '_' station '_' namaTestor '.dat'];
[~, nama] = GarjasGetNama(id_user);
nama = strrep(nama,' ','_');

if exist(namaFile, 'file') == 2
    try
        fid=fopen(namaFile, 'a+');
        fprintf(fid, '\n %12s %12s %12s \n', [id_user ' ' nama ' '  int2str(nilai) ' ']);
        fclose(fid);
        flag = true;
    catch
        warning('Cannot write a file!');
        flag = false;
    end
else
    try
        fid=fopen(namaFile, 'w');
        fprintf(fid, [ header1 ' ' header2 ' ' header3 ' ' '\n' ]);
        fprintf(fid, '%12s %12s %12s \n', [id_user ' ' nama ' '  int2str(nilai) ' ']);
        fclose(fid);
        flag = true;
    catch
        warning('Cannot write a file!');
        flag = false;
    end
end