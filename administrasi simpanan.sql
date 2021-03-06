CREATE TABLE administrasisimpanan (kode VARCHAR (20), rekening VARCHAR (20), jumlah DOUBLE, rekeningsimpanan VARCHAR(20), rekeningpendapatan VARCHAR(20), username VARCHAR (20), datetime DATETIME) TYPE = MyISAM;
ALTER TABLE golongantabungan ADD rekeningadministrasi CHAR(20) AFTER RekeningBunga;
ALTER TABLE tabungan CHANGE LastUpdateBunga LastUpdateBunga DATE DEFAULT "0000-00-00";
ALTER TABLE tabungan CHANGE JumlahBlokir JumlahBlokir DOUBLE DEFAULT "0";
ALTER TABLE tabungan CHANGE Close Close CHAR(1);
ALTER TABLE tabungan CHANGE PDL PDL CHAR(4);
ALTER TABLE tabungan CHANGE KODE KODE CHAR(10);
ALTER TABLE tabungan CHANGE StatusBlokir StatusBlokir CHAR(1);
ALTER TABLE tabungan CHANGE GOLONGANTABUNGAN GOLONGANTABUNGAN CHAR(6);
ALTER TABLE tabungan CHANGE TGL TGL DATE;
ALTER TABLE administrasisimpanan ADD tgl DATE AFTER rekeningpendapatan;