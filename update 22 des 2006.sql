/*Update biaya administrasi tabungan beserta pos jurnalnya */
ALTER TABLE tabungan ADD biayaadministrasi DOUBLE AFTER nopdl;
ALTER TABLE tabungan ADD kasbank CHAR(20)  AFTER biayaadministrasi;
ALTER TABLE tabungan ADD rekeningpendapatan CHAR(25)  AFTER kasbank;