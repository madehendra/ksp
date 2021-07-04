ALTER TABLE debitur ADD caraperhitungan CHAR(1)  AFTER DendaTelatBayar;
ALTER TABLE nomorfaktur CHANGE Kode Kode CHAR(20) NOT NULL;
