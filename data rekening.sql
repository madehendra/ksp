# HeidiSQL Dump 
#
# --------------------------------------------------------
# Host:                 localhost
# Database:             semaya
# Server version:       5.0.16-nt
# Server OS:            Win32
# max_allowed_packet:   1048576
# HeidiSQL version:     3.0 RC4 Revision: 334
# --------------------------------------------------------

/*!40100 SET CHARACTER SET latin1;*/


#
# Table structure for table 'setupbilyet'
#

CREATE TABLE /*!32312 IF NOT EXISTS*/ `setupbilyet` (
  `XRekening` double NOT NULL default '0',
  `YRekening` double NOT NULL default '0',
  `WRekening` double NOT NULL default '0',
  `XNama` double NOT NULL default '0',
  `YNama` double NOT NULL default '0',
  `WNama` double NOT NULL default '0',
  `XAlamat` double NOT NULL default '0',
  `YAlamat` double NOT NULL default '0',
  `WAlamat` double NOT NULL default '0',
  `XJumlah` double NOT NULL default '0',
  `YJumlah` double NOT NULL default '0',
  `WJumlah` double NOT NULL default '0',
  `XTerbilang` double NOT NULL default '0',
  `YTerbilang` double NOT NULL default '0',
  `WTerbilang` double NOT NULL default '0',
  `XLama` double NOT NULL default '0',
  `YLama` double NOT NULL default '0',
  `WLama` double NOT NULL default '0',
  `XValuta` double NOT NULL default '0',
  `YValuta` double NOT NULL default '0',
  `WValuta` double NOT NULL default '0',
  `XTempo` double NOT NULL default '0',
  `YTempo` double NOT NULL default '0',
  `WTempo` double NOT NULL default '0',
  `XBunga` double NOT NULL default '0',
  `YBunga` double NOT NULL default '0',
  `WBunga` double NOT NULL default '0',
  `XDirut` double NOT NULL default '0',
  `YDirut` double NOT NULL default '0',
  `WDirut` double NOT NULL default '0',
  `XKasir` double NOT NULL default '0',
  `YKasir` double NOT NULL default '0',
  `WKasir` double NOT NULL default '0',
  `Atas` double NOT NULL default '0',
  `Kiri` double NOT NULL default '0',
  `Tinggi` double NOT NULL default '0',
  `Lebar` double NOT NULL default '0',
  `xTerbilangSB` double NOT NULL default '0',
  `yTerbilangSB` double NOT NULL default '0',
  `wTerbilangSB` double NOT NULL default '0',
  `xTglCetak` double NOT NULL default '0',
  `yTglCetak` double NOT NULL default '0',
  `wTglCetak` double NOT NULL default '0',
  `xNominalBack` double default NULL,
  `yNominalBack` double default NULL,
  `wNominalBack` double default NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Dumping data for table 'setupbilyet'
#

/*!40000 ALTER TABLE setupbilyet DISABLE KEYS;*/
LOCK TABLES setupbilyet WRITE;
REPLACE INTO setupbilyet (XRekening, YRekening, WRekening, XNama, YNama, WNama, XAlamat, YAlamat, WAlamat, XJumlah, YJumlah, WJumlah, XTerbilang, YTerbilang, WTerbilang, XLama, YLama, WLama, XValuta, YValuta, WValuta, XTempo, YTempo, WTempo, XBunga, YBunga, WBunga, XDirut, YDirut, WDirut, XKasir, YKasir, WKasir, Atas, Kiri, Tinggi, Lebar, xTerbilangSB, yTerbilangSB, wTerbilangSB, xTglCetak, yTglCetak, wTglCetak, xNominalBack, yNominalBack, wNominalBack) VALUES ('140','10','35','27','28','80','27','31','80','150','60','40','30','75','80','50','60','30','15','60','30','85','60','30','120','60','30','140','85','50','0','0','0','0','0','114','190','0','0','0','120','85','80','55','75','80');
UNLOCK TABLES;
/*!40000 ALTER TABLE setupbilyet ENABLE KEYS;*/
