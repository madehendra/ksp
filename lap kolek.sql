ALTER TABLE debitur ADD wajibpokok DOUBLE AFTER JatuhTempo;
ALTER TABLE deposito ADD jumlahperpanjangan INT UNSIGNED DEFAULT "0" AFTER Lama;
ALTER TABLE mutasideposito ADD pajak DOUBLE AFTER Jumlah;
ALTER TABLE deposito CHANGE Status Status CHAR(1)  DEFAULT "0";
ALTER TABLE deposito CHANGE Kode Kode CHAR(10); 
ALTER TABLE deposito CHANGE GolonganDeposito GolonganDeposito CHAR(6); 
ALTER TABLE deposito CHANGE NominalDeposito NominalDeposito DOUBLE;
ALTER TABLE deposito CHANGE PersentaseFinalti PersentaseFinalti DOUBLE;
ALTER TABLE deposito CHANGE StatusBlokir StatusBlokir CHAR(1);
ALTER TABLE deposito CHANGE SukuBunga SukuBunga DOUBLE;
ALTER TABLE deposito CHANGE LastUpdateBunga LastUpdateBunga DATE;
ALTER TABLE deposito CHANGE StatusPostingPokok StatusPostingPokok CHAR(1);
ALTER TABLE deposito CHANGE LastPerpanjangan LastPerpanjangan DATE;
/* Khusus untuk rekening Pak Santiawan D1*/
UPDATE deposito SET LastPerpanjangan= '2006-09-03', jumlahperpanjangan= 1 WHERE ID=1;
INSERT INTO mutasideposito (Posting, ID, Faktur, Rekening, Jumlah, pajak, UserName, DateTime, DK, RekeningJurnal, Tgl, Status, KodeMutasi) VALUES ('0', NULL, 'DP012005100300000001', '01.D1.000001.01', 20000, NULL, 'TELLER1', NULL, NULL, NULL, '2005-10-03', NULL, '3');
/*Tambahan untuk jumlah simpanan yg mengendap*/
CREATE TABLE simpananmengendap (rekening VARCHAR (20), tahun VARCHAR (5), bulan VARCHAR (5), jumlah DOUBLE);
ALTER TABLE simpananmengendap ADD kode VARCHAR(15)  AFTER jumlah;
ALTER TABLE simpananmengendap ADD golongantabungan VARCHAR(5); 
CREATE TABLE postingakhirhari (tgl DATE, username VARCHAR (20), datetime DATETIME);
