# HeidiSQL Dump 
#
# --------------------------------------------------------
# Host:                 localhost
# Database:             ganesa
# Server version:       5.0.16-nt
# Server OS:            Win32
# max_allowed_packet:   1048576
# HeidiSQL version:     3.0 RC4 Revision: 334
# --------------------------------------------------------

/*!40100 SET CHARACTER SET latin1;*/


#
# Table structure for table 'saldorekening'
#

CREATE TABLE /*!32312 IF NOT EXISTS*/ `saldorekening` (
  `ID` double NOT NULL auto_increment,
  `Cabang` char(2) NOT NULL default '',
  `Rekening` char(20) NOT NULL default '',
  `AwalTahun` double default '0',
  `Awal` double default '0',
  `Debet` double default '0',
  `Kredit` double default '0',
  `Akhir` double default '0',
  PRIMARY KEY  (`ID`),
  KEY `Cabang` (`Cabang`,`Rekening`),
  KEY `Rekening` (`Rekening`,`Cabang`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Dumping data for table 'saldorekening'
#

/*!40000 ALTER TABLE saldorekening DISABLE KEYS;*/
LOCK TABLES saldorekening WRITE;
REPLACE INTO saldorekening (ID, Cabang, Rekening, AwalTahun, Awal, Debet, Kredit, Akhir) VALUES ('7','01','1.100.01','42500889','42500889','0','0','0'),
	('8','01','1.200.01','32027373','32027373','0','0','0'),
	('9','01','1.300.10','165060943','165060943','0','0','0'),
	('10','01','1.300.03','228236890','228236890','0','0','0'),
	('11','01','1.300.09','91370371','91370371','0','0','0'),
	('12','01','1.500.04','17248000','17248000','0','0','0'),
	('13','01','2.100.01.01','2381104','2381104','0','0','0'),
	('16','01','2.100.01','-277836978','-277836978','0','0','0'),
	('17','01','2.200.01','-26500000','-26500000','0','0','0'),
	('18','01','3.200','-118368591','-118368591','0','0','0'),
	('19','01','3.100.01','-32600000','-32600000','0','0','0'),
	('20','01','3.100.02','-4870000','-4870000','0','0','0'),
	('21','01','3.100.03','-118650000','-118650000','0','0','0');
UNLOCK TABLES;
/*!40000 ALTER TABLE saldorekening ENABLE KEYS;*/
