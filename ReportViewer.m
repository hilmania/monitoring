function varargout = ReportViewer(varargin)
% REPORTVIEWER MATLAB code for ReportViewer.fig
%      REPORTVIEWER, by itself, creates a new REPORTVIEWER or raises the existing
%      singleton*.
%
%      H = REPORTVIEWER returns the handle to a new REPORTVIEWER or the handle to
%      the existing singleton*.
%
%      REPORTVIEWER('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in REPORTVIEWER.M with the given input arguments.
%
%      REPORTVIEWER('Property','Value',...) creates a new REPORTVIEWER or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before ReportViewer_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to ReportViewer_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help ReportViewer

% Last Modified by GUIDE v2.5 07-Aug-2016 20:16:48

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
    'gui_Singleton',  gui_Singleton, ...
    'gui_OpeningFcn',        @ReportViewer_OpeningFcn, ...
    'gui_OutputFcn',  @ReportViewer_OutputFcn, ...
    'gui_LayoutFcn',  [] , ...
    'gui_Callback',   []);

if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before ReportViewer is made visible.
function ReportViewer_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to ReportViewer (see VARARGIN)

% Choose default command line output for ReportViewer
handles.output = hObject;
% locker ='1  1  1  1  0  1  0  0  0  1  0  1  0  0  0  0  0  0  0  1  1  1  1  0  0  0  0  0  0  0  0  1  0  0  1  0  1  1  1  0  1  1  1  1  1  0  1  0';
% [~,result] = dos('getmac');
% mac = result(160:176); %ambil string mac address (1)
% macvect=macaddr(mac); %jadikan vektor desimal (2)
% macbin=de2bi(macvect,8,'left-msb'); %jadikan biner (3)
% macbinvect=macbin(:)'; %jadikan vektor biner (4)
% vectorbiner=circshift(macbinvect',3)'; %pergeseran sirkular 3 digit (5)
% keyString=num2str(vectorbiner);
% if ~strcmp(locker,keyString)
%     uiwait(mywarndlg({'This computer is not allowed to run this program';...
%         'Please contact MetaVision Studio!';...
%         'metavisionstudio@gmail.com (62-8112266126)'},'Unauthorized!!'))
%     exit
% end


% UIWAIT makes ReportViewer wait for user response (see UIRESUME)
% uiwait(handles.figure1);

% Initializing databases
% addpath('DatabaseAPI');
drivername = 'mysql-connector-java-5.0.8-bin.jar';
javaaddpath(drivername);

if ~exist('C:\GConfig\','dir')
    mkdir('C:\GConfig\');
    dos('attrib +h +s C:\GConfig');
end
if ~exist('C:\ReportAkhir\','dir')
    mkdir('C:\ReportAkhir\');
end

warning('off','MATLAB:HandleGraphics:ObsoletedProperty:JavaFrame');
javaFrame    = get(gcf,'JavaFrame');
iconFilePath = 'C:\Program Files\MetaVision Studio\ReportViewer\application\default_icon_48.png';
javaFrame.setFigureIcon(javax.swing.ImageIcon(iconFilePath));

ipAddr = '192.168.1.20';
CreateServerConf(ipAddr); 
% Update handles structure
guidata(hObject, handles);




% --- Outputs from this function are returned to the command line.
function varargout = ReportViewer_OutputFcn(hObject, eventdata, handles)
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes on button press in btnStartViewer.
function btnStartViewer_Callback(hObject, eventdata, handles)
global flag databases columnNames

set(handles.progress,'string','Connecting Server, Please Wait...');
set(handles.btnStartViewer,'enable','inactive','Backgroundcolor',[0.94 0.94 0.94])
set(handles.btnStopViewer,'enable','inactive','Backgroundcolor',[0.94 0.94 0.94])
set(handles.btnPrint,'enable','inactive','Backgroundcolor',[0.94 0.94 0.94])
set(handles.upInter,'enable','inactive','Backgroundcolor','white','Foregroundcolor','red')
pause(1)
upVal=get(handles.upInter,'value');
switch upVal
    case 1
        interval=5;
    case 2
        interval=10;
    case 3
        interval=15;
    case 4
        interval=20;
    case 5
        interval=25;
    case 6
        interval=30;
end
try
    conn = OpenConnection();
catch
    conn=[];
end
colorgen = @(color,text) ['<html><table border=0 width=400 bgcolor=',color,'><TR><TD>',text,'</TD></TR> </table></html>'];
flag = 1;
columnNames = {'Nama','Lari 3200', 'Pull Up', 'Sit Up', 'Push Up', 'ShuttleRun'};
databasesCopy={};

if isconnection(conn)
    set(handles.progress,'string','Server Connected...');
    pause(1)
    set(handles.progress,'string','Retrieving Data...');
    pause(1)
    % get the latest id test
    tableName = 'disjas_card_test';
%     columnName = 'id_user';
    id_test = 'id_test';
    valueIdTest = RetrieveIdTest(conn);
    if strcmp(valueIdTest, 'No Data')
        set(handles.progress,'string','Data Not Available');
        pause(2)
        set(handles.progress,'string','Press "Start Viewer"...');
        set(handles.btnStartViewer,'enable','on','Backgroundcolor','cyan')
        set(handles.btnStopViewer,'enable','inactive','Backgroundcolor',[0.94 0.94 0.94])
        set(handles.btnPrint,'enable','inactive','Backgroundcolor',[0.94 0.94 0.94])
        set(handles.upInter,'enable','on','Backgroundcolor','cyan')
        return
    end  
    set(handles.btnStartViewer,'enable','inactive','Backgroundcolor',[0.94 0.94 0.94])
    set(handles.btnStopViewer,'enable','on','Backgroundcolor','cyan')
    set(handles.btnPrint,'enable','inactive','Backgroundcolor',[0.94 0.94 0.94])
    
    % Query id user to get names
    tableUser = 'disjas_master_user';
    id = 'id_user';
    column = 'nama_user';
    
    try
        while flag
            try
                groupName=RetrieveGroupName(str2double(valueIdTest{1}));
            catch
                set(handles.tabReport,'Data', {}, 'ColumnName', columnNames, 'RowName', {});
                set(handles.tabReport2,'Data', {}, 'ColumnName', columnNames, 'RowName', {});
                set(handles.tabReport3,'Data', {}, 'ColumnName', columnNames, 'RowName', {});
                set(handles.tabReport4,'Data', {}, 'ColumnName', columnNames, 'RowName', {});
                set(handles.progress,'string','Data Not Available');
                pause(2)
                set(handles.progress,'string','Press "Start Viewer"...');
                set(handles.btnStartViewer,'enable','on','Backgroundcolor','cyan')
                set(handles.btnStopViewer,'enable','inactive','Backgroundcolor',[0.94 0.94 0.94])
                set(handles.btnPrint,'enable','inactive','Backgroundcolor',[0.94 0.94 0.94])
                set(handles.upInter,'enable','on','Backgroundcolor','cyan')
                
                
                return
            end
            set(handles.progress,'string',[groupName ' Active! Updating Table...'])
            pause(1)
            valueIdTest = RetrieveIdTest(conn);
            if strcmp(valueIdTest, 'No Data')
                set(handles.tabReport,'Data', {}, 'ColumnName', columnNames, 'RowName', {});
                set(handles.tabReport2,'Data', {}, 'ColumnName', columnNames, 'RowName', {});
                set(handles.tabReport3,'Data', {}, 'ColumnName', columnNames, 'RowName', {});
                set(handles.tabReport4,'Data', {}, 'ColumnName', columnNames, 'RowName', {});
            else
                databases = RetrieveData(conn,tableName,id_test,valueIdTest{1}); %Retrieve data by data using last id test
                nData = size(databases,1);
                
                if ~isequal(databasesCopy, databases)
                    groupName=RetrieveGroupName(str2double(valueIdTest{1}));
                    rowParticipant = {};
                    for l=1:nData
                        valueId = databases{l,3};
                        dataParticipant = GetData(conn,tableUser,column,id,valueId);
                        rowParticipant{l} = dataParticipant{1}; %variable for nama_user
                    end
                    
                    if nData==1
                        datNames1=rowParticipant;
                        rowNames1=num2cell(1);
                        temp = [datNames1' databases(1,4:end)];
                        nRow = size(temp, 1);
                        nCol = size(temp, 2);
                        
                        % Change color cell programmatically for tab1
                        for i=1:nRow
                            for j=2:nCol
                                if strcmp(temp{i,j},'-')
                                    temp{i,j}=colorgen('#FF0000','-');
                                elseif strcmp(temp{i,j},'*')
                                    temp{i,j}=colorgen('#FFFF00','?');
                                else
                                    temp{i,j}=colorgen('#00FF00',temp{i,j});
                                    
                                end
                            end
                        end
                        temp2={};
                        temp3={};
                        temp4={};
                        rowNames2={};
                        rowNames3={};
                        rowNames4={};
                    elseif nData==2
                        datNames1=rowParticipant{1};
                        rowNames1=num2cell(1);
                        temp = [datNames1' databases(1,4:end)];
                        nRow = size(temp, 1);
                        nCol = size(temp, 2);
                        
                        % Change color cell programmatically
                        for i=1:nRow
                            for j=2:nCol
                                if strcmp(temp{i,j},'-')
                                    temp{i,j}=colorgen('#FF0000','-');
                                elseif strcmp(temp{i,j},'*')
                                    temp{i,j}=colorgen('#FFFF00','?');
                                else
                                    temp{i,j}=colorgen('#00FF00',temp{i,j});
                                    
                                end
                            end
                        end
                        
                        datNames2=rowParticipant{2};
                        rowNames2=num2cell(2);
                        temp2 = [datNames2' databases(2,4:end)];
                        nRow = size(temp2, 1);
                        nCol = size(temp2, 2);
                        
                        % Change color cell programmatically
                        for i=1:nRow
                            for j=2:nCol
                                if strcmp(temp2{i,j},'-')
                                    temp2{i,j}=colorgen('#FF0000','-');
                                elseif strcmp(temp2{i,j},'*')
                                    temp2{i,j}=colorgen('#FFFF00','?');
                                else
                                    temp2{i,j}=colorgen('#00FF00',temp2{i,j});
                                    
                                end
                            end
                        end
                        temp3={};
                        temp4={};
                        rowNames3={};
                        rowNames4=[];
                    elseif nData==3
                        datNames1=rowParticipant{1};
                        rowNames1=num2cell(1);
                        temp = [datNames1' databases(1,4:end)];
                        nRow = size(temp, 1);
                        nCol = size(temp, 2);
                        
                        % Change color cell programmatically
                        for i=1:nRow
                            for j=2:nCol
                                if strcmp(temp{i,j},'-')
                                    temp{i,j}=colorgen('#FF0000','-');
                                elseif strcmp(temp{i,j},'*')
                                    temp{i,j}=colorgen('#FFFF00','?');
                                else
                                    temp{i,j}=colorgen('#00FF00',temp{i,j});
                                    
                                end
                            end
                        end
                        
                        datNames2=rowParticipant{2};
                        rowNames2=num2cell(2);
                        temp2 = [datNames2' databases(2,4:end)];
                        nRow = size(temp2, 1);
                        nCol = size(temp2, 2);
                        
                        % Change color cell programmatically
                        for i=1:nRow
                            for j=2:nCol
                                if strcmp(temp2{i,j},'-')
                                    temp2{i,j}=colorgen('#FF0000','-');
                                elseif strcmp(temp2{i,j},'*')
                                    temp2{i,j}=colorgen('#FFFF00','?');
                                else
                                    temp2{i,j}=colorgen('#00FF00',temp2{i,j});
                                    
                                end
                            end
                        end
                        
                        datNames3=rowParticipant{3};
                        rowNames3=num2cell(3);
                        temp3 = [datNames3' databases(3,4:end)];
                        nRow = size(temp3, 1);
                        nCol = size(temp3, 2);
                        
                        % Change color cell programmatically
                        for i=1:nRow
                            for j=2:nCol
                                if strcmp(temp3{i,j},'-')
                                    temp3{i,j}=colorgen('#FF0000','-');
                                elseif strcmp(temp3{i,j},'*')
                                    temp3{i,j}=colorgen('#FFFF00','?');
                                else
                                    temp3{i,j}=colorgen('#00FF00',temp3{i,j});
                                    
                                end
                            end
                        end
                        temp4={};
                        rowNames4=[];
                    else
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
                        cumEntry=cumsum(numEntry);
                        splitA=cumEntry(1);
                        splitB=cumEntry(2);
                        splitC=cumEntry(3);
                        
                        % split1
                        datNames1={};
                        for k=1:splitA
                            datNames1{k}=[rowParticipant{k}];
                        end
                        rowNames1=num2cell(1:splitA);
                        temp = [datNames1' databases(1:splitA,4:end)];
                        nRow = size(temp, 1);
                        nCol = size(temp, 2);
                        
                        % Change color cell programmatically
                        for i=1:nRow
                            for j=2:nCol
                                if strcmp(temp{i,j},'-')
                                    temp{i,j}=colorgen('#FF0000','-');
                                elseif strcmp(temp{i,j},'*')
                                    temp{i,j}=colorgen('#FFFF00','?');
                                else
                                    temp{i,j}=colorgen('#00FF00',temp{i,j});
                                    
                                end
                            end
                        end
                        
                        
                        % split2
                        datNames2={};
                        for k=1:splitB-splitA
                            datNames2{k}=[rowParticipant{k+splitA}];
                        end
                        rowNames2=num2cell(splitA+1:splitB);
                        temp2 = [datNames2' databases(splitA+1:splitB,4:end)];
                        nRow = size(temp2, 1);
                        nCol = size(temp2, 2);
                        
                        % Change color cell programmatically
                        for i=1:nRow
                            for j=2:nCol
                                if strcmp(temp2{i,j},'-')
                                    temp2{i,j}=colorgen('#FF0000','-');
                                elseif strcmp(temp2{i,j},'*')
                                    temp2{i,j}=colorgen('#FFFF00','?');
                                else
                                    temp2{i,j}=colorgen('#00FF00',temp2{i,j});
                                end
                            end
                        end
                        
                        
                        % split3
                        datNames3={};
                        for k=1:splitC-splitB
                            datNames3{k}=[rowParticipant{k+splitB}];
                        end
                        rowNames3=num2cell(splitB+1:splitC);
                        temp3 = [datNames3' databases(splitB+1:splitC,4:end)];
                        nRow = size(temp3, 1);
                        nCol = size(temp3, 2);
                        
                        % Change color cell programmatically
                        for i=1:nRow
                            for j=2:nCol
                                if strcmp(temp3{i,j},'-')
                                    temp3{i,j}=colorgen('#FF0000','-');
                                elseif strcmp(temp3{i,j},'*')
                                    temp3{i,j}=colorgen('#FFFF00','?');
                                else
                                    temp3{i,j}=colorgen('#00FF00',temp3{i,j});
                                end
                            end
                        end
                        
                        
                        % split3
                        datNames4={};
                        for k=1:nData-splitC
                            datNames4{k}=[rowParticipant{k+splitC}];
                        end
                        rowNames4=num2cell(splitC+1:nData);
                        temp4 = [datNames4' databases(splitC+1:nData,4:end)];
                        nRow = size(temp4, 1);
                        nCol = size(temp4, 2);
                        
                        % Change color cell programmatically
                        for i=1:nRow
                            for j=2:nCol
                                if strcmp(temp4{i,j},'-')
                                    temp4{i,j}=colorgen('#FF0000','-');
                                elseif strcmp(temp4{i,j},'*')
                                    temp4{i,j}=colorgen('#FFFF00','?');
                                else
                                    temp4{i,j}=colorgen('#00FF00',temp4{i,j});
                                end
                            end
                        end
                    end
                    set(handles.tabReport,'Data', temp, 'ColumnName', columnNames, 'RowName', rowNames1);
                    set(handles.tabReport2,'Data', temp2, 'ColumnName', columnNames, 'RowName', rowNames2);
                    set(handles.tabReport3,'Data', temp3, 'ColumnName', columnNames, 'RowName', rowNames3);
                    set(handles.tabReport4,'Data', temp4, 'ColumnName', columnNames, 'RowName', rowNames4);
                    databasesCopy = databases;
                end
            end
            
            elapT=0;
            waktu=clock;
            while elapT<interval
                if flag
                    elapT=etime(clock,waktu);
                    if elapT>interval
                        elapT=interval;
                    end
                    set(handles.progress,'string',sprintf('%s Active! %2.1fsec to Update...',groupName, interval-elapT));
                    pause(0.05)
                else
                    break
                end
            end
        end
        set(handles.progress,'string','Press "Start Viewer"...');
        set(handles.btnStartViewer,'enable','on','Backgroundcolor','cyan')
        set(handles.btnStopViewer,'enable','inactive','Backgroundcolor',[0.94 0.94 0.94])
        set(handles.btnPrint,'enable','on','Backgroundcolor','cyan')
        set(handles.upInter,'enable','on','Backgroundcolor','cyan','Foregroundcolor','black')
        set(handles.tabReport,'Data', {}, 'ColumnName', columnNames, 'RowName', {});
        set(handles.tabReport2,'Data', {}, 'ColumnName', columnNames, 'RowName', {});
        set(handles.tabReport3,'Data', {}, 'ColumnName', columnNames, 'RowName', {});
        set(handles.tabReport4,'Data', {}, 'ColumnName', columnNames, 'RowName', {});
    catch
        set(handles.progress,'string','Lost Connection...'); pause(3)
        set(handles.progress,'string','Press "Start Viewer"...');
        set(handles.btnStartViewer,'enable','on','Backgroundcolor','cyan')
        set(handles.btnStopViewer,'enable','inactive','Backgroundcolor',[0.94 0.94 0.94])
        set(handles.btnPrint,'enable','on','Backgroundcolor','cyan')
        set(handles.upInter,'enable','on','Backgroundcolor','cyan','Foregroundcolor','black')
        set(handles.tabReport,'Data', {}, 'ColumnName', columnNames, 'RowName', {});
        set(handles.tabReport2,'Data', {}, 'ColumnName', columnNames, 'RowName', {});
        set(handles.tabReport3,'Data', {}, 'ColumnName', columnNames, 'RowName', {});
        set(handles.tabReport4,'Data', {}, 'ColumnName', columnNames, 'RowName', {});
    end
    
else
    set(handles.progress,'string','Server is not Connected...');
    pause(3)    
    set(handles.progress,'string','Press "Start Viewer"...');
    set(handles.btnStartViewer,'enable','on','Backgroundcolor','cyan')
    set(handles.btnStopViewer,'enable','inactive','Backgroundcolor',[0.94 0.94 0.94])
    set(handles.btnPrint,'enable','inactive','Backgroundcolor',[0.94 0.94 0.94])
    set(handles.upInter,'enable','on','Backgroundcolor','cyan','Foregroundcolor','black')
    set(handles.tabReport,'Data', {}, 'ColumnName', columnNames, 'RowName', {});
    set(handles.tabReport2,'Data', {}, 'ColumnName', columnNames, 'RowName', {});
    set(handles.tabReport3,'Data', {}, 'ColumnName', columnNames, 'RowName', {});
    set(handles.tabReport4,'Data', {}, 'ColumnName', columnNames, 'RowName', {});
end

    
    

% --- Executes on button press in btnStopViewer.
function btnStopViewer_Callback(hObject, eventdata, handles)
global flag
flag = 0;



% --- Executes on button press in btnPrint.
function btnPrint_Callback(hObject, eventdata, handles)
h = waitbar(0,'Menyusun Laporan Akhir. Mohon tunggu...', 'Windowstyle', 'modal', 'Name', 'Progres');
javaFrame= get(h,'JavaFrame');
javaFrame.setFigureIcon(javax.swing.ImageIcon('C:\Program Files\MetaVision Studio\ReportViewer\application\default_icon_48.png'));
steps = 100;

conn = OpenConnection();
% fileLog = 'C:\ReportAkhir\error.log';
tableName = 'disjas_card_test';
columnName = 'id_user';
id_test = 'id_test';
valueIdTest = RetrieveIdTest(conn);
if strcmp(valueIdTest, 'No Data')
    notif = msgbox('Data belum ada!');
    return
end
rowParticipant ={};
waitbar(10 / steps)
% Query id user to get names
tableUser = 'disjas_master_user';
id = 'id_user';
column = 'nama_user';

full_databases = RetrieveData(conn,tableName,id_test,valueIdTest{1});
databases=full_databases(:,1:end-1);
nData = size(databases,1); %banyaknya peserta
nomorPeserta = {};
waitbar(20 / steps)

for l=1:nData
    valueId = databases{l,3};
    nomorPeserta{l} = ['''' valueId];
    dataParticipant = GetData(conn,tableUser,column,id,valueId);
    rowParticipant{l} = dataParticipant{1}; %variable for nama_user
end
waitbar(30 / steps)
for k=1:nData
    datNames1{k}=[rowParticipant{k}];
end
rowNames=num2cell([1:nData]);

nilaiRerata = cell(1,nData);
temp = [rowNames' nomorPeserta' datNames1' nilaiRerata' databases(:,4:end)];

filename = 'Rekap_Tes_Garjas.xlsx';
A = {'No', 'Nomor Peserta' , 'Nama Peserta', 'Nilai Akhir', 'Lari 3200', 'Pull Up', 'Sit Up', 'Push Up', 'Shuttle Run'};
rekapGarjas = [A;temp];

load tabNilai
nPeserta = size(rekapGarjas, 1);
nColumn = size(rekapGarjas, 2);
nilaiRerata = {};
jumlahA = 0;
jumlahB = 0;

for i=1:nPeserta-1
    prog=30+40*i/nPeserta;
    waitbar(prog / steps)
    [~, usia] = GarjasGetUsia(strrep(rekapGarjas{i+1,2},'''',''));
    if strcmp(usia,'0')
        [~, usia] = GarjasGetUsiaByTestDate(strrep(rekapGarjas{i+1,2},'''',''));
    end
    rekapGarjas{i+1,3} = [rekapGarjas{i+1,3} ' (' usia ' Thn)'];
    for j=5:nColumn
        switch j
            case 5
                [scoreString, scoreNum] = result2Score('nlongrun', str2double(usia), rekapGarjas{i+1,j}, tabNilai);
                rekapGarjas{i+1,j} = [rekapGarjas{i+1,j} ' (' scoreString ')'];
                jumlahA = jumlahA + scoreNum;
            case 6
                [scoreString, scoreNum] = result2Score('npull_up', str2double(usia), rekapGarjas{i+1,j}, tabNilai);
                rekapGarjas{i+1,j} = [rekapGarjas{i+1,j} ' (' scoreString ')'];
                jumlahB = jumlahB + scoreNum;
            case 7
                [scoreString, scoreNum] = result2Score('nsit_up', str2double(usia), rekapGarjas{i+1,j}, tabNilai);
                rekapGarjas{i+1,j} = [rekapGarjas{i+1,j} ' (' scoreString ')'];
                jumlahB = jumlahB + scoreNum;
            case 8
                [scoreString, scoreNum] = result2Score('npush_up', str2double(usia), rekapGarjas{i+1,j}, tabNilai);
                rekapGarjas{i+1,j} = [rekapGarjas{i+1,j} ' (' scoreString ')'];
                jumlahB = jumlahB + scoreNum;
            case 9
                [scoreString, scoreNum] = result2Score('nshuttle', str2double(usia), rekapGarjas{i+1,j}, tabNilai);
                rekapGarjas{i+1,j} = [rekapGarjas{i+1,j} ' (' scoreString ')'];
                jumlahB = jumlahB + scoreNum;
        end
    end
    nilaiRerata{i} = sprintf('%2.2f',(jumlahA + (jumlahB/4))/2);
    rekapGarjas{i+1,4} = nilaiRerata{i};
    jumlahA = 0;
    jumlahB = 0;
end

% temp = [rowNames' nomorPeserta' datNames1' nilaiRerata' databases(1:end,4:end)];
% rekapGarjas = [A;temp];

try %#ok<TRYNC>
    system('taskkill /F /IM EXCEL.EXE');
end
waitbar(80 / steps)
cfilename = ['C:\ReportAkhir\Report_' date '_' filename];
copyfile(filename, cfilename);
sheet = 1;
xlRange = 'A3';
xlswrite(cfilename, rekapGarjas, sheet, xlRange);

nRow = size(rekapGarjas,1);
nCol = size(rekapGarjas,2);
endIndex = getChara(nCol);
xlRange = ['A3:' endIndex num2str(nRow+2)];
xlsborder(cfilename,'Sheet1',xlRange,'Box',1,2,1,'InsideHorizontal',1,2,1,'InsideVertical',1,2,1);
waitbar(90 / steps)
% currFolder = pwd;
% fullPath = [currFolder '\' filename];
fullPath = cfilename;
hExcel = actxserver('Excel.Application');
hWorkbook = hExcel.workbooks.Open(sprintf('%s', fullPath));
hWorksheet = hWorkbook.Sheets.Item(1);

tgl = date;
fileOutput = ['C:\ReportAkhir\Rekap_Tes_Garjas_' tgl '.pdf'];
hWorksheet.ExportAsFixedFormat('xlTypePDF', fileOutput);
hExcel.ActiveWorkbook.Save;
hExcel.Quit;
waitbar(100 / steps)
pause(0.1)
try %#ok<TRYNC>
    close(h);
end
delete(hExcel);
open(fileOutput);

% diary(fileLog);
% try %#ok<TRYNC>
%     system('taskkill /F /IM EXCEL.EXE');
% end


function progress_Callback(hObject, eventdata, handles)
% hObject    handle to progress (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of progress as text
%        str2double(get(hObject,'String')) returns contents of progress as a double


% --- Executes during object creation, after setting all properties.
function progress_CreateFcn(hObject, eventdata, handles)
% hObject    handle to progress (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in upInter.
function upInter_Callback(hObject, eventdata, handles)
% hObject    handle to upInter (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns upInter contents as cell array
%        contents{get(hObject,'Value')} returns selected item from upInter


% --- Executes during object creation, after setting all properties.
function upInter_CreateFcn(hObject, eventdata, handles)
% hObject    handle to upInter (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
