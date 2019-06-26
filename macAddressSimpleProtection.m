clc
clear
close all
[status,result] = dos('getmac');
mac = result(160:176); %ambil string mac address (1)
macvect=macaddr(mac); %jadikan vektor desimal (2)
macbin=de2bi(macvect,8,'left-msb'); %jadikan biner (3)
macbinvect=macbin(:)'; %jadikan vektor biner (4)
vectorbiner=circshift(macbinvect',3)'; %pergeseran sirkular 3 digit (5)
key = num2str(vectorbiner)
%vektor biner ini yang akan jadi kunci (6)

% Tahapan membuat kunci/locker
% 1. baca macaddresnpada pc yang akan diinstal
% 2. jadikan vektor desimal
% 3. jadikan elemen vektor desimal tersebut matrik biner row-wise untuk setiap elemen
% 4. jadikan matrik biner row-wise tersebut vektor biner dengan pembacaan collumn-wise (efek interlace)
% 5. lakukan pergeseran sirkular 3 digit untuk mendapatkan kunci bagi komputer yang bersangkutan
% 6. gunakan kunci tersebut sebagai unlocker pada opening function program
% station

% cara menggunakan kunci
% 1. dapatkan kunci pada komputer dengan cara diatas (Tahapan membuat
% kunci/locker) untuk digunakan pada opening function
% 2. pada opening function sisipkan code untuk mendapatkan kunci
% 3. pada opening function bandingkan kunci pada step 1 dan step 2, jika sama lanjutkan (do nothing), jika tidak sama (munculkan warning "not authorize" dan return
