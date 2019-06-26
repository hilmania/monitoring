function [scoreString, scoreNum]=result2Score(station, usia, hasil, tabNilai)
listUsia=18:57;
if usia<18
    usia=18;
elseif usia >57
    usia=57;
end

ix=find(listUsia==usia);
catUsia=ceil(ix/4);
colPointer=catUsia+1;

try
    if strcmp(station,'nlongrun')%nilai lari===============================
        semcol=find(hasil==':');
        if isempty(semcol)
            waktu=str2double(hasil);
        else
            waktu=60*str2double(hasil(1:semcol-1)) + str2double(hasil(semcol+1:end));
        end
        if waktu==0
            scoreString='0';
            scoreNum=0;
        elseif waktu < 657
            scoreString='100';
            scoreNum=str2double(scoreString);
        elseif waktu > 2230
            scoreString='0';
            scoreNum=str2double(scoreString);
        else
            stringWaktu=tabNilai(:,1);
            listWaktu=zeros(151,1);
            for k=1:151
                listWaktu(k)=str2double(stringWaktu{k});
            end
            if ~isnan(waktu)
                deltaTime=listWaktu-waktu;
                [~, ix]=min(abs(deltaTime));
                if waktu - listWaktu(ix)>0
                    rowPointer=ix+1;
                else
                    rowPointer=ix;
                end
                scoreString=tabNilai{rowPointer,colPointer};
                scoreNum=str2double(scoreString);
            else
                scoreString='0';
                scoreNum=0;
            end
        end
    elseif strcmp(station,'nshuttle')%nilai shutle run=====================
        waktu=str2double(hasil);
        if waktu < 15.9
            scoreString='100';
            scoreNum=str2double(scoreString);
        elseif waktu > 30.9
            scoreString='0';
            scoreNum=str2double(scoreString);
        else
            stringWaktu=tabNilai(:,15);
            listWaktu=zeros(151,1);
            for k=1:151
                listWaktu(k)=str2double(stringWaktu{k});
            end
            rowPointer= listWaktu==waktu;
            scoreString=tabNilai{rowPointer,colPointer};
            scoreNum=str2double(scoreString);
        end
    elseif strcmp(station,'npush_up')%nilai push up========================
        counter=str2double(hasil);
        if counter > 43
            scoreString='100';
            scoreNum=str2double(scoreString);
        elseif counter < 1
            scoreString='0';
            scoreNum=str2double(scoreString);
        else
            stringCount=tabNilai(:,14);
            listCount=zeros(151,1);
            for k=1:151
                listCount(k)=str2double(stringCount{k});
            end
            rowPointer= listCount==counter;
            scoreString=tabNilai{rowPointer,colPointer};
            scoreNum=str2double(scoreString);
        end
    elseif strcmp(station,'nsit_up')%nilai sit up==========================
        counter=str2double(hasil);
        if counter > 45
            scoreString='100';
            scoreNum=str2double(scoreString);
        elseif counter < 1
            scoreString='0';
            scoreNum=str2double(scoreString);
        else
            stringCount=tabNilai(:,13);
            listCount=zeros(151,1);
            for k=1:151
                listCount(k)=str2double(stringCount{k});
            end
            rowPointer= listCount==counter;
            scoreString=tabNilai{rowPointer,colPointer};
            scoreNum=str2double(scoreString);
        end
    elseif strcmp(station,'npull_up')%nilai pull up========================
        counter=str2double(hasil);
        if counter > 18
            scoreString='100';
            scoreNum=str2double(scoreString);
        elseif counter < 1
            scoreString='0';
            scoreNum=str2double(scoreString);
        else
            stringCount=tabNilai(:,12);
            listCount=zeros(151,1);
            for k=1:151
                listCount(k)=str2double(stringCount{k});
            end
            rowPointer= listCount==counter;
            scoreString=tabNilai{rowPointer,colPointer};
            scoreNum=str2double(scoreString);
        end
    end
catch
    scoreString='0';
    scoreNum=0;
end