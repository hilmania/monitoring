nData=153
numEntry=zeros(1,4);
siklus=floor(nData/4);
if siklus>=1
    numEntry=repmat(siklus,1,4);
    sisa=nData-4*siklus;
    for k=1:sisa
        numEntry(k)=numEntry(k)+1;
    end
else
    for k=1:nData
        numEntry(k)=numEntry(k)+1;
    end
end

numEntry
sum(numEntry)