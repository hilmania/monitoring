clc
clear
close all

% [num, text] = xlsread('Norma Nilai.xlsx');
% for b=1:151
%     for k=1:13
%         text{b,k+1}=num2str(num(b,k));
%     end
% end
% for k=1:151
%     dat=text{k,1};
%     semcol=find(dat==':');
%     sec=60*str2num(dat(1:semcol-1)) + str2num(dat(semcol+1:end));
%     text{k,1}=num2str(sec);
% end
% tabNilai=text;
% save tabNilai tabNilai
% %==========================================================


load tabNilai
usia=18;
station='nshuttle'; % nlongrun nsit_up npull_up npush_up nshuttle
hasil='20';


[scoreString, scoreNum]=result2Score(station, usia, hasil, tabNilai)
nilai = {};

load rekapHasil
load tabNilai
nPeserta = size(rekapGarjas, 1);
nColumn = size(rekapGarjas, 2);
nilaiRerata = {};
jumlahA = 0;
jumlahB = 0;

for i=1:nPeserta-1
    [~, usia] = GarjasGetUsia(rekapGarjas{i+1,2});
    rekapGarjas{i+1,2} = [rekapGarjas{i+1,2} ' (' usia ')'];
    
    for j=1:nColumn
        switch j
            case 4
                [scoreString, scoreNum] = result2Score('nlongrun', str2num(usia), rekapGarjas{i+1,j}, tabNilai);
                rekapGarjas{i+1,j} = [rekapGarjas{i+1,j} ' (' scoreString ')'];
                jumlahA = jumlahA + scoreNum
            case 5
                [scoreString, scoreNum] = result2Score('npull_up', str2num(usia), rekapGarjas{i+1,j}, tabNilai);
                rekapGarjas{i+1,j} = [rekapGarjas{i+1,j} ' (' scoreString ')'];
                jumlahB = jumlahB + scoreNum
            case 6
                [scoreString, scoreNum] = result2Score('nsit_up', str2num(usia), rekapGarjas{i+1,j}, tabNilai);
                rekapGarjas{i+1,j} = [rekapGarjas{i+1,j} ' (' scoreString ')'];
                jumlahB = jumlahB + scoreNum
            case 7
                [scoreString, scoreNum] = result2Score('npush_up', str2num(usia), rekapGarjas{i+1,j}, tabNilai);
                rekapGarjas{i+1,j} = [rekapGarjas{i+1,j} ' (' scoreString ')'];
                jumlahB = jumlahB + scoreNum
            case 8
                [scoreString, scoreNum] = result2Score('nshuttle', str2num(usia), rekapGarjas{i+1,j}, tabNilai);
                rekapGarjas{i+1,j} = [rekapGarjas{i+1,j} ' (' scoreString ')'];
                jumlahB = jumlahB + scoreNum
        end
    end
    nilaiRerata{i+1} = sprintf('%2.2f',(jumlahA + (jumlahB/4))/2);
    jumlahA = 0;
    jumlahB = 0;
end




% listUsia=18:57;
% if usia<18
%     usia=18;
% elseif usia >57
%     usia=57;
% end
% ix=find(listUsia==usia);
% catUsia=ceil(ix/4);
% colPointer=catUsia+1;
% if strcmp(station,'nlongrun')%nilai lari===========
%     semcol=find(hasil==':');
%     waktu=60*str2double(hasil(1:semcol-1)) + str2double(hasil(semcol+1:end));
%     if waktu < 657
%         nilai=100;
%     elseif waktu > 2230
%         nilai=0;
%     else
%         stringWaktu=tabNilai(:,1);
%         listWaktu=zeros(151,1);
%         for k=1:151
%             listWaktu(k)=str2double(stringWaktu{k});
%         end
%         deltaTime=listWaktu-waktu;
%         [val, ix]=min(abs(deltaTime));
%         if waktu - listWaktu(ix)>0
%             rowPointer=ix+1;
%         else
%             rowPointer=ix;
%         end
%         nilai=tabNilai{rowPointer,colPointer};
%     end
% elseif strcmp(station,'nshuttle')
%     waktu=str2double(hasil);
%     if waktu < 15.9
%         nilai=100;
%     elseif waktu > 30.9
%         nilai=0;
%     else
%         stringWaktu=tabNilai(:,15);
%         listWaktu=zeros(151,1);
%         for k=1:151
%             listWaktu(k)=str2double(stringWaktu{k});
%         end
%         rowPointer=find(listWaktu==waktu);
%         nilai=tabNilai{rowPointer,colPointer};
%     end
% elseif strcmp(station,'npush_up')
%     counter=str2double(hasil);
%     if counter > 43
%         nilai=100;
%     elseif counter < 1
%         nilai=0;
%     else
%         stringCount=tabNilai(:,14);
%         listCount=zeros(151,1);
%         for k=1:151
%             listCount(k)=str2double(stringCount{k});
%         end
%         rowPointer=find(listCount==counter);
%         nilai=tabNilai{rowPointer,colPointer};
%     end
% elseif strcmp(station,'nsit_up')
%     counter=str2double(hasil);
%     if counter > 45
%         nilai=100;
%     elseif counter < 1
%         nilai=0;
%     else
%         stringCount=tabNilai(:,13);
%         listCount=zeros(151,1);
%         for k=1:151
%             listCount(k)=str2double(stringCount{k});
%         end
%         rowPointer=find(listCount==counter);
%         nilai=tabNilai{rowPointer,colPointer};
%     end
% elseif strcmp(station,'npull_up')
%     counter=str2double(hasil);
%     if counter > 18
%         nilai=100;
%     elseif counter < 1
%         nilai=0;
%     else
%         stringCount=tabNilai(:,12);
%         listCount=zeros(151,1);
%         for k=1:151
%             listCount(k)=str2double(stringCount{k});
%         end
%         rowPointer=find(listCount==counter);
%         nilai=tabNilai{rowPointer,colPointer};
%     end
% end
% nilai








