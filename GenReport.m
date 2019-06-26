function GenReport(dataReport)
% Init Data Report
masterReport = fileread('template.html');

% Find and replace header information like $nama / $nomor dll
exprHeader = {'\<[$]nama','\<[$]nomor','\<[$]ttl','\<[$]pdk','\<[$]daerah'};
replHeader = {dataReport(1).info.nama,dataReport(1).info.nomor,dataReport(1).info.ttl,...
    dataReport(1).info.pdk,dataReport(1).info.daerah};

newReport = regexprep(masterReport,exprHeader,replHeader);

% Find and replace information about score, assert the right score, and sum
% all the score
abjad=char(97:114);
num=1:4;
fild={};
ix=1;
for k=1:numel(abjad)
    for m=1:4
        fild{ix}=['''' '\' '<' '[$]' abjad(k) num2str(num(m)) ''''];
        ix=ix+1;
    end
end

str='exprScore ={';
for k=1:72
    str=cat(2,str,fild{k},',');
end
str(end)='}';
str=[str ';'];
eval(str);

ix=1;
for k=1:numel(abjad)
    for m=1:4
        repfild{ix}=dataReport(ix).score;
        ix=ix+1; 
    end
end

repstr='replScore ={';
for k=1:72
    repstr=cat(2,repstr,'''',repfild{k},'''',',');
end
repstr(end)='}';
repstr=[repstr ';'];
eval(repstr);

newReport = regexprep(newReport,exprScore,replScore);
%--------------------------------------------------------------------------

% Find and replace footer information
exprFooter = {'\<[$]berat','\<[$]tinggi','\<[$]nilai','\<[$]kat','\<[$]tempat','\<[$]testor','\<[$]rekomendasi',...
    '\<[$]pengamatan','\<[$]sikap'};
replFooter = {dataReport(1).info.berat,dataReport(1).info.tinggi,dataReport(1).info.nilai,...
    dataReport(1).info.kat,dataReport(1).info.tempat,dataReport(1).info.testor, dataReport(1).info.rekomendasi,...
    dataReport(1).info.pengamatan,dataReport(1).info.sikap};
newReport = regexprep(newReport,exprFooter,replFooter);

% Find and insert general and detil image
image = 'img';
numImage = 1:24;
imageData = {};
idx=1;
for m=1:24
    imageData{idx}=['''' image num2str(numImage(m)) ''''];
    idx=idx+1;
end

strImage='exprImage ={';
for k=1:24
    strImage=cat(2,strImage,imageData{k},',');
end
strImage(end)='}';
strImage=[strImage ';'];
eval(strImage);

idx2=1;
for m=1:24
    replImageData{idx2}=dataReport(idx2).image;
    idx2=idx2+1; 
end

replStrImage='replImage ={';
for k=1:24
    replStrImage=cat(2,replStrImage,'''',replImageData{k},'''',',');
end
replStrImage(end)='}';
replStrImage=[replStrImage ';'];
eval(replStrImage);

newReport = regexprep(newReport,exprImage,replImage);
%--------------------------------------------------------------------------

% Write generated report to a HTML file
generatedReport = fopen('report.html','w');
fwrite(generatedReport,newReport);
fclose(generatedReport);