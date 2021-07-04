VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm aMainmenu 
   BackColor       =   &H8000000C&
   Caption         =   "USAHA SIMPAN PINJAM DESA"
   ClientHeight    =   7110
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   13125
   Icon            =   "aMainmenu.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   255
      Top             =   1695
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pcCancel 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      Picture         =   "aMainmenu.frx":030A
      ScaleHeight     =   285
      ScaleWidth      =   13065
      TabIndex        =   1
      Top             =   1155
      Visible         =   0   'False
      Width           =   13125
   End
   Begin VB.PictureBox pcExit 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      Picture         =   "aMainmenu.frx":0499
      ScaleHeight     =   285
      ScaleWidth      =   13065
      TabIndex        =   0
      Top             =   810
      Visible         =   0   'False
      Width           =   13125
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   2220
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aMainmenu.frx":052F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aMainmenu.frx":0849
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aMainmenu.frx":0B63
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aMainmenu.frx":0E7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aMainmenu.frx":1197
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aMainmenu.frx":14B1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   13125
      _ExtentX        =   23151
      _ExtentY        =   1429
      ButtonWidth     =   1879
      ButtonHeight    =   1429
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "E&xit"
            Key             =   "Close"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Log Off.."
            Key             =   "LogOff"
            Object.ToolTipText     =   "Log Off"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Printer"
            Key             =   "Printer"
            Object.ToolTipText     =   "Setup Printer"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Help"
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Calculator"
            Key             =   "Calculator"
            Object.ToolTipText     =   "Calculator"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&About.."
            Key             =   "About"
            Object.ToolTipText     =   "About"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   6795
      Width           =   13125
      _ExtentX        =   23151
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9701
            MinWidth        =   9701
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "4:31 AM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "05-Jul-21"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuInputrekeing 
         Caption         =   "Proses Input Rekening"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMasterCabang 
         Caption         =   "Master &Cabang"
      End
      Begin VB.Menu mnuMasterAkuntansi 
         Caption         =   "Master &Akuntansi"
         Begin VB.Menu mnuMasterRekening 
            Caption         =   "Master &Rekening"
         End
         Begin VB.Menu mnuMasterSaldoAwalRekening 
            Caption         =   "Master &Saldo Awal Rekening"
         End
      End
      Begin VB.Menu mnuSptFile1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMasterAgama 
         Caption         =   "Master A&gama"
      End
      Begin VB.Menu mnuMasterPekerjaan 
         Caption         =   "Master &Pekerjaan"
      End
      Begin VB.Menu mnuMasterDaerah 
         Caption         =   "Master &Daerah"
      End
      Begin VB.Menu mnuMasterRegisterNasabah 
         Caption         =   "Master &Register Anggota"
      End
      Begin VB.Menu mnuSpt3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMasterTabungan 
         Caption         =   "Master &Simpanan"
         Begin VB.Menu mnuMstPDL 
            Caption         =   "Master PDL"
         End
         Begin VB.Menu mnuGolonganTabungan 
            Caption         =   "Golongan Simpanan"
         End
         Begin VB.Menu mnuSukuBunga 
            Caption         =   "Suku Bunga Progressif"
         End
         Begin VB.Menu mnuKodeTransaksi 
            Caption         =   "Kode Transaksi"
         End
         Begin VB.Menu mnuKonfigurasiTabungan 
            Caption         =   "Konfigurasi Simpanan"
         End
         Begin VB.Menu mnuSaldoAwalTabungan 
            Caption         =   "Saldo Awal Simpanan"
         End
      End
      Begin VB.Menu mnuMasterDeposito 
         Caption         =   "Master &Deposito"
      End
      Begin VB.Menu mnuMasterKredit 
         Caption         =   "Master &Pinjaman"
         Begin VB.Menu mnuGolonganKredit 
            Caption         =   "Golongan Pinjaman"
         End
         Begin VB.Menu mnuSetupJaminan 
            Caption         =   "Setup Jaminan"
         End
         Begin VB.Menu mnuAccountOfficer 
            Caption         =   "Account Officer"
         End
      End
      Begin VB.Menu mnuSptFile2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogOff 
         Caption         =   "Log Off..."
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Transaksi"
      Index           =   1
      Begin VB.Menu mnuTrTeller 
         Caption         =   "Teller"
      End
      Begin VB.Menu mnuTrTabungan 
         Caption         =   "Simpanan"
         Begin VB.Menu mnuTrPembukaanRekeningTabungan 
            Caption         =   "Pembukaan Rekening Simpanan"
         End
         Begin VB.Menu mnuTrBlokirTabungan 
            Caption         =   "Buka / Blokir Simpanan"
         End
         Begin VB.Menu mnuKoreksiMutasitabungan 
            Caption         =   "Koreksi Mutasi Simpanan"
         End
         Begin VB.Menu mnuTrHapusKoreksiMutasiTabungan 
            Caption         =   "Hapus Mutasi Simpanan"
         End
         Begin VB.Menu mnuTrTutuptabungan 
            Caption         =   "Tutup Rekening Simpanan"
         End
         Begin VB.Menu sptAdministrasi 
            Caption         =   "-"
         End
         Begin VB.Menu mnuProsesBiayaAdministrasi 
            Caption         =   "Proses Administrasi"
         End
         Begin VB.Menu mnuPembatalanProsesAdministrasi 
            Caption         =   "Pembatalan Proses Administrasi"
         End
      End
      Begin VB.Menu mnuTrDeposito 
         Caption         =   "Deposito"
         Begin VB.Menu mnuTrPembuakaanDeposito 
            Caption         =   "Pembukaan Deposito"
         End
         Begin VB.Menu mnuCetakBilyet 
            Caption         =   "Cetak Bilyet"
         End
         Begin VB.Menu mnuPembatalanBilyet 
            Caption         =   "Pembatalan Bilyet"
         End
         Begin VB.Menu mnuTrBlokirBungaDeposito 
            Caption         =   "Blokir Deposito"
         End
         Begin VB.Menu mnutrHapusMutasiDep 
            Caption         =   "Hapus Mutasi Deposito"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuTrSpt 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuTrCetakValidasiDeposito 
            Caption         =   "Cetak Validasi Deposito"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuTrKredit 
         Caption         =   "Pinjaman"
         Begin VB.Menu mnuTrPenganjuanKredit 
            Caption         =   "Pengajuan Pinjaman"
         End
         Begin VB.Menu mnuTrRealisasiKredit 
            Caption         =   "Realisasi Pinjaman"
         End
         Begin VB.Menu mnuCetakPinjaman 
            Caption         =   "Cetak Pinjaman"
         End
         Begin VB.Menu mnuTrKoreksiAngsuran 
            Caption         =   "Koreksi Angsuran Pinjaman"
         End
         Begin VB.Menu mnutrHapusANgsKredit 
            Caption         =   "Hapus Pencairan / Angsuran Pinjaman"
         End
         Begin VB.Menu sptSaldoAwalKredit 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSaldoAwalKredit 
            Caption         =   "Saldo Awal Kredit"
         End
         Begin VB.Menu mnuIlustrasiPinjaman 
            Caption         =   "Ilustrasi Pinjaman"
         End
         Begin VB.Menu mnuEditAO 
            Caption         =   "Edit AO"
         End
      End
      Begin VB.Menu mnuSepJU 
         Caption         =   "-"
      End
      Begin VB.Menu mnutrjurnalUmum 
         Caption         =   "Jurnal Umum"
      End
      Begin VB.Menu mnuPencetakan 
         Caption         =   "Pencetakan"
         Visible         =   0   'False
         Begin VB.Menu test 
            Caption         =   "Test"
         End
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Laporan"
      Index           =   2
      Begin VB.Menu mnuRptRegisternasabah 
         Caption         =   "Laporan Register Nasabah"
      End
      Begin VB.Menu mnuRptRekapRegisternasabah 
         Caption         =   "Rekapitulasi Register Nasabah"
      End
      Begin VB.Menu mnuRptMutasiTeller 
         Caption         =   "Mutasi harian Teller"
      End
      Begin VB.Menu mnuDaftarAnggota 
         Caption         =   "Daftar Anggota"
      End
      Begin VB.Menu mnUSepDaftarnasabah 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRptTabungan 
         Caption         =   "Simpanan"
         Begin VB.Menu mnuRptSaldoTabungan 
            Caption         =   "Laporan Saldo Simpanan"
         End
         Begin VB.Menu mnuRptBukutabungan 
            Caption         =   "Laporan Buku Simpanan"
         End
         Begin VB.Menu mnuRptmutasitabunganHarian 
            Caption         =   "Laporan Mutasi Simpanan Harian"
         End
         Begin VB.Menu mnuRptBungaTabungan 
            Caption         =   "Laporan Bunga Simpanan"
         End
         Begin VB.Menu rptRekapitulasiTabunganKeseluruhanPDL 
            Caption         =   "Laporan Rekapitulasi Simpanan Keseluruhan"
         End
         Begin VB.Menu sptLaporanSimpanan 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLaporanSimpananPokok 
            Caption         =   "Laporan Simpanan Pokok"
         End
         Begin VB.Menu mnuLaporanSimpananWajib 
            Caption         =   "Laporan Simpanan Wajib"
         End
      End
      Begin VB.Menu mnuRptDeposito 
         Caption         =   "Deposito"
         Begin VB.Menu mnuLaporanDaftarDeposan 
            Caption         =   "Laporan Daftar Deposan"
         End
         Begin VB.Menu mnuMutasiDeposito 
            Caption         =   "Laporan Mutasi Deposito"
         End
         Begin VB.Menu mnuRptDepositojatuhtempo 
            Caption         =   "Laporan Deposito Jatuh Tempo"
         End
         Begin VB.Menu mnuLaporanTurunBungaDeposito 
            Caption         =   "Laporan Turun Bunga Deposito"
         End
         Begin VB.Menu mnuLaporanRekapitulasiDeposito 
            Caption         =   "Laporan Rekapitulasi Deposito"
         End
         Begin VB.Menu mnuLaporanKartuBunga 
            Caption         =   "Laporan Kartu Bunga"
         End
      End
      Begin VB.Menu mnuRptKredit 
         Caption         =   "Pinjaman"
         Begin VB.Menu mnuRptPengajuanKredit 
            Caption         =   "Laporan Pengajuan Pinjaman"
         End
         Begin VB.Menu mnuRptRealisasiKredit 
            Caption         =   "Laporan Reliasasi Pinjaman"
         End
         Begin VB.Menu mnuRptMutasiHarianKredit 
            Caption         =   "Laporan Mutasi Harian Pinjaman"
         End
         Begin VB.Menu mnuRptJadwalAngsuran 
            Caption         =   "Laporan Jadwal Angsuran"
         End
         Begin VB.Menu mnuRptBukuAngsuran 
            Caption         =   "Laporan Buku Angsuran"
         End
         Begin VB.Menu mnuRptTagihanKredit 
            Caption         =   "Laporan Daftar Tagihan Pinjaman"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuTunggakanKredit 
            Caption         =   "Laporan Tunggakan Pinjaman"
            Visible         =   0   'False
         End
         Begin VB.Menu MnuRptbakiDebet 
            Caption         =   "Laporan Baki Debet"
         End
         Begin VB.Menu mnuRptKreditjatuhTempo 
            Caption         =   "Laporan Pinjaman Jatuh Tempo"
         End
         Begin VB.Menu mnuRptTurunBunga 
            Caption         =   "Laporan Turun Bunga"
         End
         Begin VB.Menu mnuLaporanKreditLunas 
            Caption         =   "Laporan Pinjaman Yang Sudah Lunas"
         End
         Begin VB.Menu mnuLaporanKreditPerTanggal 
            Caption         =   "Laporan Pinjaman Per Tanggal"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuLaporanYgNunggakBulanan 
            Caption         =   "Laporan Tagihan Pinjaman"
         End
         Begin VB.Menu mnuLaporanAgunan 
            Caption         =   "Laporan Agunan"
         End
         Begin VB.Menu mnuDebtCollector 
            Caption         =   "Debt Collector"
            Begin VB.Menu mnuRealisasiPinjaman 
               Caption         =   "Realisasi Pinjaman"
            End
            Begin VB.Menu mnuAngsuranPerDebt 
               Caption         =   "Angsuran per Debt"
            End
         End
      End
      Begin VB.Menu mnUSepAKNt 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRptAkuntansi 
         Caption         =   "Akuntansi"
         Begin VB.Menu mnuRptDaftarRekening 
            Caption         =   "Laporan Daftar Rekening"
         End
         Begin VB.Menu mnuRptJurnalHarian 
            Caption         =   "Laporan Jurnal Harian"
         End
         Begin VB.Menu mnurekapJurnalharian 
            Caption         =   "Laporan Rekap Jurnal Harian"
         End
         Begin VB.Menu mnuRptBukuBesar 
            Caption         =   "Laporan Buku Besar"
         End
         Begin VB.Menu mnuLaporanArusKas 
            Caption         =   "Laporan Arus Kas"
         End
         Begin VB.Menu mnuSepLBR 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRptLabaRugi 
            Caption         =   "Laporan Laba / Rugi"
         End
         Begin VB.Menu mnuNeracaLajur 
            Caption         =   "Neraca Percobaan..."
         End
         Begin VB.Menu mnuRptNeraca 
            Caption         =   "Laporan Neraca"
         End
      End
      Begin VB.Menu sptSHU 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLaporanSHU 
         Caption         =   "SHU"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWajibPokok 
         Caption         =   "Kolektibilitas..."
      End
      Begin VB.Menu mnuLaporanRatioRatioFinancial 
         Caption         =   "Laporan Ratio - Ratio Financial"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Utility"
      Index           =   3
      Begin VB.Menu mnuUpdate 
         Caption         =   "Update"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPostingBungaDanPokokDeposito 
         Caption         =   "Posting Bunga dan Pokok Deposito"
         Begin VB.Menu mnPostinAwal 
            Caption         =   "Posting Awal Hari (Bunga + Pokok Deposito)"
         End
         Begin VB.Menu mnuPembatalanProsesAwalHari 
            Caption         =   "Pembatalan Proses Awal Hari"
         End
         Begin VB.Menu mnuPotingAkhirhari 
            Caption         =   "Posting Akhir Hari (Perpajangan Deposito ARO)"
         End
      End
      Begin VB.Menu mnuPengendapan 
         Caption         =   "Posting Akhir Hari (Pengendapan Saldo Simpanan Pokok & Wajib)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPostingBungaSimpananReguler 
         Caption         =   "Posting Bunga Simpanan Reguler (Mengendap Satu Bulan)"
         Begin VB.Menu mnuPostingBungaTabungan 
            Caption         =   "Posting Bunga Simpanan"
         End
         Begin VB.Menu mnuBatalPostingBungaTabungan 
            Caption         =   "Batal Posting Bunga Simpanan"
         End
      End
      Begin VB.Menu mnuPostingBungaSimpananHarian 
         Caption         =   "Posting Bunga Simpanan Harian"
         Begin VB.Menu mnuPostingBungaHarian 
            Caption         =   "Posting Bunga Harian"
         End
         Begin VB.Menu mnuPostingAwalBulan 
            Caption         =   "Posting Bunga Pas Akhir Bulan"
         End
      End
      Begin VB.Menu munCariRegister 
         Caption         =   "Cari Register Nasabah"
      End
      Begin VB.Menu mnuSPTPosting 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPostingBukuBesar 
         Caption         =   "Posting Buku Besar"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTimbrah 
         Caption         =   "Timbrah"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPostingCikarSedana 
         Caption         =   "Cikar Sedana"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUpdatePosting 
         Caption         =   "Update Posting"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Setup"
      Index           =   4
      Begin VB.Menu MnuPassword 
         Caption         =   "&Create User Name"
      End
      Begin VB.Menu MnuMenuLevel 
         Caption         =   "&Setup Menu Level"
      End
      Begin VB.Menu MnuChangePassword 
         Caption         =   "&Change Password"
      End
      Begin VB.Menu SetupSep 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCfgSetPrinter 
         Caption         =   "Setup Printer"
      End
      Begin VB.Menu MnuSetupInfoPerusahaan 
         Caption         =   "Informasi Perusahaan"
      End
      Begin VB.Menu mnuWallpaper 
         Caption         =   "Wallpaper"
      End
      Begin VB.Menu mnuSPT 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKonfigurasiBukuTabungan 
         Caption         =   "Konfigurasi Buku Simpanan"
      End
      Begin VB.Menu mnuSptBukuTabungan 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetupKonfigurasiKodeTransaksi 
         Caption         =   "Konfigurasi Kode Transaksi"
      End
      Begin VB.Menu mnuSetupKonfigurasiKasTeller 
         Caption         =   "Konfigurasi Kas  Teller"
      End
      Begin VB.Menu mnuSetupTeller 
         Caption         =   "Setup Menu Teller"
      End
      Begin VB.Menu mnuSetupKonfigurasiBilyetDeposito 
         Caption         =   "Konfigurasi Bilyet Deposito"
      End
      Begin VB.Menu mnuKonfAutoJurnal 
         Caption         =   "Konfigurasi Rekening Auto Jurnal"
      End
      Begin VB.Menu mnuSepPer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStpPeriodeTabungan 
         Caption         =   "Setup Periode Bunga Simpanan"
      End
      Begin VB.Menu mnuSetupPeriode 
         Caption         =   "Setup Periode Akuntansi"
      End
      Begin VB.Menu mnuSetupAnggota 
         Caption         =   "Setup Anggota"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSetupKeterangandanJabatanCetakanNeraca 
         Caption         =   "Setup Keterangan dan Jabatan Cetakan Neraca"
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Sisa &Hasil Usaha"
      Index           =   5
      Visible         =   0   'False
      Begin VB.Menu mnuSetupSisaHasilUsaha 
         Caption         =   "Setup Sisa Hasil Usaha"
      End
      Begin VB.Menu mnuSisaHasilUsaha 
         Caption         =   "Sisa Hasil Usaha"
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Windows"
      Index           =   6
      WindowList      =   -1  'True
      Begin VB.Menu MnuWindows 
         Caption         =   "Tile Horizontally"
         Index           =   0
      End
      Begin VB.Menu MnuWindows 
         Caption         =   "Tile Vertically"
         Index           =   1
      End
      Begin VB.Menu MnuWindows 
         Caption         =   "Cascade"
         Index           =   2
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Help"
      Index           =   7
      Visible         =   0   'False
      Begin VB.Menu MnuHelpContents 
         Caption         =   "Contents..."
      End
      Begin VB.Menu MnuHelpIndex 
         Caption         =   "Index..."
      End
      Begin VB.Menu MnuHelpSearch 
         Caption         =   "Search..."
         Visible         =   0   'False
      End
      Begin VB.Menu SepHelp2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuAboutHelp 
         Caption         =   "About"
      End
      Begin VB.Menu HelpTest 
         Caption         =   "Test"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuBsmUtility 
      Caption         =   "BSM_Utility"
      Visible         =   0   'False
      Begin VB.Menu mnuTrUpdate 
         Caption         =   "Update"
      End
      Begin VB.Menu mnuTutupBuku 
         Caption         =   "Tutup Buku"
      End
      Begin VB.Menu mnuBatalTutupBuku 
         Caption         =   "Batal Tutup Buku"
      End
      Begin VB.Menu mnuPenghapusanBungaTabungan 
         Caption         =   "Penghapusan Bunga Tabungan"
      End
      Begin VB.Menu mnuUpdateSystemRegister 
         Caption         =   "Update System Register"
      End
      Begin VB.Menu bsmSpt 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRekapitulasiPencairanPokokBunga 
         Caption         =   "Rekapitulasi Pencairan Pokok/Bunga Deposito"
      End
      Begin VB.Menu mnuHapusRecordBukuBesar 
         Caption         =   "Hapus Record Buku Besar (Mutasi Deposito)"
      End
      Begin VB.Menu mnuHapusSeluruhRecordBungaDeposito 
         Caption         =   "Hapus Seluruh Record di Table Bunga Deposito"
      End
      Begin VB.Menu sptlain 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUpdateGolonganDeposito 
         Caption         =   "Update Golongan Deposito"
      End
   End
   Begin VB.Menu mnuPosting1 
      Caption         =   "Posting"
      Visible         =   0   'False
      Begin VB.Menu mnuPosting 
         Caption         =   "Posting"
      End
   End
   Begin VB.Menu About 
      Caption         =   "&About"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "aMainmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lFirst As Boolean
Dim cKode As String
Dim objMenu As New CodeSuiteLibrary.Menu
Dim objData As New CodeSuiteLibrary.data
Dim dbData As New ADODB.Recordset
Dim cIPNumber As String
Dim cNamaDatabase As String
Dim cNamaDSN As String
Dim cPort As String
Dim cKeySecret As String
Dim cMYODBCPATH As String
Dim cMYODBCFile As String
Dim cModePelunasanPiutang As String

Private Sub HelpTest_Click()
  Load frmTest
  frmTest.Show
End Sub

Private Sub MDIForm_Activate()
  If lFirst Then
'    If CekTrialVersionAktif = False Then
      Me.Caption = aCfg(msNama)
      lFirst = False
      
      'SendKeys "{S}{Enter}{s}{Enter}{Enter}"
      If Not objMenu.GetPassword(cKode, Me, GetDSN) Then
        End
      End If
      

  
      mnuLogOff.Caption = "&Log Off " & Trim(objMenu.FullName) & "..."
      Toolbar1.Buttons(2).ToolTipText = mnuLogOff.Caption
      Toolbar1.Buttons(6).ToolTipText = "About " & GetAppDescription
      nUserLevel = objMenu.UserLevel
      cUserID = objMenu.UserID
      cusername = objMenu.UserName
      cFullName = objMenu.FullName
  
      SaveRegistry reg_FullName, objMenu.FullName
      SaveRegistry reg_UserLevel, objMenu.UserLevel
      SaveRegistry reg_UserName, objMenu.UserName
      SaveRegistry reg_UserID, objMenu.UserID
  
      StatusBar1.Panels(1).Text = "USER: " & objMenu.UserName & " _SERVER: " & UCase(GetRegistry(reg_IP)) & " _DATABASE: " & UCase(GetRegistry(reg_Database))
      StatusBar1.Panels(2).Text = GetAppDescription
      Me.Picture = LoadPicture(GetPicture(GetRegistry(reg_Wallpaper)))
  
      Set dbData = objData.Browse(GetDSN, "username", "kasteller", "username", sisAssign, objMenu.UserName)
      If dbData.RecordCount > 0 Then
        cKasTeller = GetNull(dbData!KasTeller, "")
      End If
      
      'Otentikasi
'      If authKey(aCfg(msSerialKey), MBSerialNumber) = False Then
'        If CheckTrial(objData, 100) = True Then
'          Load frmAbout
'          frmAbout.Show vbModal
'        End If
'      End If
    
'    Else
'      MsgBox "Maaf. Program dibatasi pemakaiannya. Untuk bisa digunakan lebih lanjut silahkan hubungi Bali Surya Media (0361-228198)"
'      End
'    End If
  End If
End Sub

Function CheckTrial(ByVal obj As CodeSuiteLibrary.data, ByVal nTrial As Double) As Boolean
Dim db As New ADODB.Recordset
Dim nRecords As Double


  nTrial = 0
  Set db = obj.Browse(GetDSN, "bukubesar", "faktur")
  If Not db.eof Then
    nRecords = nTrial + db.RecordCount
  End If
         
  If nRecords > nTrial Then
    MsgBox "Maaf" & vbCrLf & "Program ini adalah versi Trial. Batas pemakaian dibatasi, dan Anda sudah memakainya dengan maksimal" & vbCrLf & "Terimakasih sudah mencoba menggunakan prgram ini" & vbCrLf & "Jika anda tertarik untuk memakai dan membeli program ini silahkan hubungi alamat email seperti yg sudah tertera dalam menu Help :: About"
    CheckTrial = True
  End If

End Function

Private Sub MDIForm_Load()
Dim lSave As Boolean
Dim cUID As String
Dim cPwd As String


  GetMyODBCFile cMYODBCFile
  GetIPNumber cIPNumber, cNamaDatabase, cNamaDSN, cPort, cKeySecret, cModePelunasanPiutang
  
  
  cUID = "kode"
  cPwd = "FullMoon"
  CreateDSN cNamaDSN, cIPNumber, cNamaDatabase, cUID, cPwd, cPort, cMYODBCPATH, cMYODBCFile
  
  
  SaveRegistry reg_ServerUID, cUID
  SaveRegistry reg_ServerPWD, cPwd
'  SaveRegistry reg_KeySecret, UCase(cKeySecret)
  
  'CreateDSN cNamaDSN, cIPNumber, cNamaDatabase, cUID, cPwd, cPort
  SaveRegistry reg_ServerUID, cUID
  SaveRegistry reg_ServerPWD, cPwd
  lFirst = True
  cKode = ""
  InitConnection
  Me.Picture = LoadPicture(GetPicture(GetRegistry(reg_Wallpaper)))
 
End Sub

Sub GetMyODBCFile(ByRef cMYODBCFile As String)
Dim cFile As String
Dim n As Double
Dim cData As String

  cFile = App.Path & "\config.ini"
  If Dir(cFile) <> "" Then
    Open cFile For Input Shared As #1
    Do While Not eof(1)
      Line Input #1, cData
      cMYODBCFile = GetData(cData, "MYODBC_FILE = ", cMYODBCFile)
    Loop
    Close #1
  End If
End Sub

Private Sub mnPostinAwal_Click()
  Load frmPostingAwalHari
  frmPostingAwalHari.Show
End Sub

Private Sub mnuAbout_Click()
  Load frmAbout
  frmAbout.Show vbModal
End Sub

Private Sub MnuAboutHelp_Click()
  Load frmAboutMe
  frmAboutMe.Show
End Sub

Private Sub mnuAccountOfficer_Click()
  Load MstAO
  MstAO.Show
End Sub

Private Sub mnuAngsuranPerDebt_Click()
  Load rptDebtKolektorAngsuran
  rptDebtKolektorAngsuran.Show
  rptDebtKolektorAngsuran.SetFocus
End Sub

Private Sub mnuBatalPostingBungaTabungan_Click()
  Load FrmBatalPostingBungaTabungan
  FrmBatalPostingBungaTabungan.Show
End Sub

Private Sub mnuBatalTutupBuku_Click()
  Load frmBatalTutupBuku
  frmBatalTutupBuku.Show
End Sub

Private Sub mnuCetakBilyet_Click()
  Load trCetakBilyet
  trCetakBilyet.Show
End Sub

Private Sub mnuCetakPinjaman_Click()
  Load trCetakPinjaman
  trCetakPinjaman.Show
End Sub

Private Sub MnuCfgSetPrinter_Click()
  CommonDialog1.ShowPrinter
End Sub

Private Sub MnuChangePassword_Click()
  objMenu.ChangePassword cKode
End Sub

Private Sub mnuDaftarAnggota_Click()
  Load rptDaftarAnggota
  rptDaftarAnggota.Show
End Sub

Private Sub mnuEditAO_Click()
  Load trEditAONasabahKredit
  trEditAONasabahKredit.Show
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuGolonganKredit_Click()
  Load MstGolonganKredit
  MstGolonganKredit.Show
End Sub

Private Sub mnuGolonganTabungan_Click()
  Load MstGolonganTabungan
  MstGolonganTabungan.Show
End Sub

Private Sub MnuHelpContents_Click()
  Load trReksadanaCalc
  trReksadanaCalc.Show vbModal
End Sub

Private Sub mnuIlustrasiPinjaman_Click()
  Load trIlustrasi
  trIlustrasi.Show
End Sub

Private Sub mnuInputrekeing_Click()
  Load aProsesTabungan
  aProsesTabungan.Show
End Sub

Private Sub mnuKodeTransaksi_Click()
  Load MstKodeTransaksi
  MstKodeTransaksi.Show
End Sub

Private Sub mnuKonfigurasiHeaderBukuTabungan_Click()
  Load frmSetupHeaderBukuTabungan
  frmSetupHeaderBukuTabungan.Show
End Sub

Private Sub mnuKodeTransaksiBaru_Click()
  Load CfgKodeTransaksi
  CfgKodeTransaksi.Show
End Sub

Private Sub mnuKonfAutoJurnal_Click()
  Load cfgAUtoJurnal
  cfgAUtoJurnal.Show
End Sub

Private Sub mnuKonfigurasiBukuTabungan_Click()
  Load cfgBukuTabungan2
  cfgBukuTabungan2.Show
End Sub

Private Sub mnuKonfigurasiTabungan_Click()
  Load cfgTabungan
  cfgTabungan.Show
End Sub

Private Sub mnuKoreksiMutasitabungan_Click()
  Load trKoreksiMutasiTabungan
  trKoreksiMutasiTabungan.Show
End Sub

Private Sub mnuLaporanAgunan_Click()
  Load rptAgunan
  rptAgunan.Show
  rptAgunan.SetFocus
End Sub

Private Sub mnuLaporanArusKas_Click()
  Load rptArusKas
  rptArusKas.Show
End Sub

Private Sub mnuLaporanDaftarDeposan_Click()
  Load RptSaldoDeposito
  RptSaldoDeposito.Show
End Sub

Private Sub mnuLaporanKartuBunga_Click()
  Load trKartuBunga
  trKartuBunga.Show
End Sub

Private Sub mnuLaporanKreditLunas_Click()
  Load rptKreditYgLunas
  rptKreditYgLunas.Show
End Sub

Private Sub mnuLaporanKreditPerTanggal_Click()
  Load rptTagihanKreditPerTgl
  rptTagihanKreditPerTgl.Show
End Sub

Private Sub mnuLaporanRatioRatioFinancial_Click()
  Load rptRatioFinancial
  rptRatioFinancial.Show
End Sub

Private Sub mnuLaporanRekapitulasiDeposito_Click()
  Load rptLaporanRekapitulasiDeposito
  rptLaporanRekapitulasiDeposito.Show
End Sub

Private Sub mnuLaporanSHUUsahaPinjaman_Click()

End Sub

Private Sub mnuLaporanSHU_Click()
  Load rptSHUPinjaman
  rptSHUPinjaman.Show
End Sub

Private Sub mnuLaporanTurunBungaDeposito_Click()
  Load rptLaporanBungaDeposito
  rptLaporanBungaDeposito.Show
End Sub

Private Sub mnuLaporanYgNunggakBulanan_Click()
  Load rptTunggakanBulanan
  rptTunggakanBulanan.Show
End Sub

Private Sub mnuLogOff_Click()
  Unload Me
  Me.Show
End Sub

Private Sub mnuMasterAgama_Click()
  Load MstAgama
  MstAgama.Show
End Sub

Private Sub mnuMasterCabang_Click()
  Load MstCabang
  MstCabang.Show
End Sub

Private Sub mnuMasterDaerah_Click()
  Load MstWilayah
  MstWilayah.Show
End Sub

Private Sub mnuMasterDeposito_Click()
  Load MstGolonganDeposito
  MstGolonganDeposito.Show
End Sub

Private Sub mnuMasterPekerjaan_Click()
  Load MstPekerjaan
  MstPekerjaan.Show
End Sub

Private Sub mnuMasterRegisterNasabah_Click()
  Load MstRegisterNasabah
  MstRegisterNasabah.Show
End Sub

Private Sub mnuMasterRekening_Click()
  Load MstRekening
  MstRekening.Show
End Sub

Private Sub mnuMasterSaldoAwalRekening_Click()
  Load MstSARekening
  MstSARekening.Show
End Sub

Private Sub MnuMenuLevel_Click()
  objMenu.SisSetMenu Me, cKode, GetDSN
End Sub

Private Sub mnuMstPDL_Click()
  Load MstPDL
  MstPDL.Show
End Sub

Private Sub mnuMutasiDeposito_Click()
  Load RptMutasiDeposito
  RptMutasiDeposito.Show
End Sub

Private Sub mnuNeracaLajur_Click()
  Load rptNeracaLajur
  rptNeracaLajur.Show
End Sub

Private Sub MnuPassword_Click()
  objMenu.AddPassword GetDSN, cKode
End Sub

Private Sub mnuPembatalanBilyet_Click()
  Load TrBatalBilyet
  TrBatalBilyet.Show
End Sub


Private Sub mnuPembatalanProsesAdministrasi_Click()
  Load trBatalAdministrasi
  trBatalAdministrasi.Show
End Sub

Private Sub mnuPembatalanProsesAwalHari_Click()
  Load frmBatalPostingAwalHari
  frmBatalPostingAwalHari.Show
End Sub

Private Sub mnuPengendapan_Click()
  Load trPostingAkhirHariPengendapan
  trPostingAkhirHariPengendapan.Show
End Sub

Private Sub mnuPenghapusanBungaTabungan_Click()
  Load trPenghapusanBungaTabungan
  trPenghapusanBungaTabungan.Show
End Sub

Private Sub mnuPosting_Click()
  Load frmPosting
  frmPosting.Show
End Sub

Private Sub mnuPostingAwalBulan_Click()
  Load cfgPostingAwalBulan
  cfgPostingAwalBulan.Show
End Sub

Private Sub mnuPostingBukuBesar_Click()
  Load trPostingBukuBesar
  trPostingBukuBesar.Show
End Sub

Private Sub mnuPostingBungaHarian_Click()
  Load trPostingBungaHarian
  trPostingBungaHarian.Show
End Sub

Private Sub mnuPostingBungaTabungan_Click()
  Load FrmPostingBungaTabungan
  FrmPostingBungaTabungan.Show
End Sub

Private Sub mnuPostingCikarSedana_Click()
  Load frmPostingCikarSedana
  frmPostingCikarSedana.Show
End Sub

Private Sub mnuPotingAkhirhari_Click()
  Load frmPostingAkhirHari
  frmPostingAkhirHari.Show
End Sub

Private Sub mnuProsesBiayaAdministrasi_Click()
  Load trAdministrasiSimpanan
  trAdministrasiSimpanan.Show
End Sub

Private Sub mnuRealisasiPinjaman_Click()
  Load rptDebtRealisasi
  rptDebtRealisasi.Show
  rptDebtRealisasi.SetFocus
End Sub

Private Sub mnuRekapitulasiPencairanPokokBunga_Click()
  Load rptRekapPencairanPokokBungaDeposito
  rptRekapPencairanPokokBungaDeposito.Show
End Sub

Private Sub mnurekapJurnalharian_Click()
  Load RptRekapitulasiJurnalHarian
  RptRekapitulasiJurnalHarian.Show
End Sub

Private Sub MnuRptbakiDebet_Click()
  Load RptBakiDebet
  RptBakiDebet.Show
End Sub

Private Sub mnuRptBukuAngsuran_Click()
  Load RptBukuAngsuran
  RptBukuAngsuran.Show
End Sub

Private Sub mnuRptBukuBesar_Click()
  Load RptBukuBesar
  RptBukuBesar.Show
End Sub

Private Sub mnuRptBukutabungan_Click()
  Load RptBukuTabungan
  RptBukuTabungan.Show
End Sub

Private Sub mnuRptBungaTabungan_Click()
  Load RptBungaTabungan
  RptBungaTabungan.Show
End Sub

Private Sub mnuRptDaftarRekening_Click()
  Load RptDaftarRekening
  RptDaftarRekening.Show
End Sub

Private Sub mnuRptDepositojatuhtempo_Click()
  Load RptJatuhTempo1
  RptJatuhTempo1.Show
End Sub

Private Sub mnuRptJadwalAngsuran_Click()
  Load RptJadwalAngsuran
  RptJadwalAngsuran.Show
End Sub

Private Sub mnuRptJurnalHarian_Click()
  Load RptJurnalHarian
  RptJurnalHarian.Show
End Sub

Private Sub mnuRptKreditjatuhTempo_Click()
  Load RptJatuhtempoKredit
  RptJatuhtempoKredit.Show
End Sub

Private Sub mnuRptLabaRugi_Click()
'  Load RptLabaRugi
'  RptLabaRugi.Show
  Load rptNewLabaRugi
  rptNewLabaRugi.Show
End Sub

Private Sub mnuRptMutasiHarianKredit_Click()
  Load RptMutasiharianKredit
  RptMutasiharianKredit.Show
End Sub

Private Sub mnuRptmutasitabunganHarian_Click()
  Load RptMutasiTabunganHarian
  RptMutasiTabunganHarian.Show
End Sub

Private Sub mnuRptMutasiTeller_Click()
  Load RptMutasiHarianTeller1
  RptMutasiHarianTeller1.Show
End Sub

Private Sub mnuRptNeraca_Click()
'  Load RptNeraca
'  RptNeraca.Show
  Load rptNeracaUpdate
  rptNeracaUpdate.Show
  rptNeracaUpdate.SetFocus
End Sub

Private Sub mnuRptPengajuanKredit_Click()
  Load RptPengajuanKredit
  RptPengajuanKredit.Show
End Sub

Private Sub mnuRptRealisasiKredit_Click()
  Load RptRealisasiKredit
  RptRealisasiKredit.Show
End Sub

Private Sub mnuRptRegisternasabah_Click()
  Load RptRegisterNasabah
  RptRegisterNasabah.Show
End Sub

Private Sub mnuRptRekapRegisternasabah_Click()
  Load RptRekapNasabah
  RptRekapNasabah.Show
End Sub

Private Sub mnuRptSaldoDeposito_Click()
End Sub

Private Sub mnuRptSaldoTabungan_Click()
'  Load RptSaldoTabungan
'  RptSaldoTabungan.Show
  
  Load rptNewSaldoTabungan
  rptNewSaldoTabungan.Show
End Sub

Private Sub mnuRptTurunBunga_Click()
  Load RptTurunBunga
  RptTurunBunga.Show
End Sub

Private Sub mnuSaldoAwalKredit_Click()
  Load trSaldoAwalKredit
  trSaldoAwalKredit.Show
End Sub

Private Sub mnuSaldoAwalTabungan_Click()
  Load MstSaldoAwalTabungan
  MstSaldoAwalTabungan.Show
End Sub

Private Sub mnuSetupAnggota_Click()
  Load mstAnggotaTetap
  mstAnggotaTetap.Show
  mstAnggotaTetap.SetFocus
End Sub

Private Sub MnuSetupInfoPerusahaan_Click()
  Load CfgInfoPerusahaan
  CfgInfoPerusahaan.Show
End Sub

Private Sub mnuSetupJaminan_Click()
  Load MstJaminan
  MstJaminan.Show
End Sub

Private Sub mnuSetupKeterangandanJabatanCetakanNeraca_Click()
  Load cfgCetakanNeracaPembuatPemeriksa
  cfgCetakanNeracaPembuatPemeriksa.Show
End Sub

Private Sub mnuSetupKonfigurasiBilyetDeposito_Click()
  Load cfgSetupBilyet
  cfgSetupBilyet.Show
End Sub

Private Sub mnuSetupKonfigurasiKasTeller_Click()
  Load CfgKasTeller
  CfgKasTeller.Show
End Sub

Private Sub mnuSetupKonfigurasiKodeTransaksi_Click()
  Load CfgKodeTransaksi
  CfgKodeTransaksi.Show
End Sub

Private Sub mnuSetupPeriode_Click()
  Load mstPeriodeAkuntansi
  mstPeriodeAkuntansi.Show
End Sub

Private Sub mnuSetupSisaHasilUsaha_Click()
  Load cfgSHU
  cfgSHU.Show
  cfgSHU.SetFocus
End Sub

Private Sub mnuSetupTeller_Click()
  Load FrmSetupTeller
  FrmSetupTeller.Show
End Sub

Private Sub mnuSisaHasilUsaha_Click()
  Load frmSHU
  frmSHU.Show
  frmSHU.SetFocus
End Sub

Private Sub mnuStpPeriodeTabungan_Click()
  Load FrmSetupPeriodeBunga
  FrmSetupPeriodeBunga.Show
End Sub

Private Sub mnuSukuBunga_Click()
  Load MstSukuBunga
  MstSukuBunga.Show
End Sub

Private Sub MnuSysInfo_Click()
  Load cfgInfo
  cfgInfo.Show
End Sub

Private Sub mnuTimbrah_Click()
  Load frmTimbrah
  frmTimbrah.Show
End Sub

Private Sub mnuTrBlokirBungaDeposito_Click()
  Load trBlokirDeposito
  trBlokirDeposito.Show
End Sub

Private Sub mnuTrBlokirTabungan_Click()
  Load trBlokirTabungan
  trBlokirTabungan.Show
End Sub

Private Sub mnutrHapusANgsKredit_Click()
  Load trHapusAngsuran
  trHapusAngsuran.Show
End Sub

Private Sub mnuTrHapusKoreksiMutasiTabungan_Click()
  Load trHapusKoreksiMutasiTabungan
  trHapusKoreksiMutasiTabungan.Show
End Sub

Private Sub mnutrHapusMutasiDep_Click()
  Load trKoreksiMutasiDeposito
  trKoreksiMutasiDeposito.Show
End Sub

Private Sub mnutrjurnalUmum_Click()
  Load trJurnalUmum
  trJurnalUmum.Show
End Sub

Private Sub mnuTrKoreksiAngsuran_Click()
  Load trKoreksiAngsuran
  trKoreksiAngsuran.Show
End Sub

Private Sub mnuTrPembuakaanDeposito_Click()
  Load trOpenDeposito
  trOpenDeposito.Show
End Sub

Private Sub mnuTrPembukaanRekeningTabungan_Click()
  Load TrOpenTabungan
  TrOpenTabungan.Show
End Sub

Private Sub mnuTrPenganjuanKredit_Click()
  Load TrPengajuanKredit
  TrPengajuanKredit.Show
End Sub

Private Sub mnuTrRealisasiKredit_Click()
  Load TrRealisasi
  TrRealisasi.Show
End Sub

Private Sub mnuTrTeller_Click()
  Load trTeller
  trTeller.Show
End Sub

Private Sub mnuTrTutuptabungan_Click()
  Load trTutupTabungan
  trTutupTabungan.Show
End Sub

Private Sub mnuTrUpdate_Click()
  Load trUpdate
  trUpdate.Show
End Sub

Private Sub mnuTutupBuku_Click()
  Load frmTutupBuku
  frmTutupBuku.Show
End Sub

Private Sub mnuUpdate_Click()
  Load frmUpdate
  frmUpdate.Show
End Sub

Private Sub mnuUpdateGolonganDeposito_Click()
  Load cfgUpdateGolonganDeposito
  cfgUpdateGolonganDeposito.Show
End Sub

Private Sub mnuUpdatePosting_Click()
  Load frmPostingNatarSari
  frmPostingNatarSari.Show
End Sub

Private Sub mnuUpdateSystemRegister_Click()
  Load frmSystemRegister
  frmSystemRegister.Show
End Sub

Private Sub mnuWajibPokok_Click()
  Load trWajibPokok
  trWajibPokok.Show
End Sub

Private Sub mnuWallpaper_Click()
  GetWallpaper
End Sub

Private Sub MnuWindows_Click(Index As Integer)
  Select Case Index
    Case 0
      Me.Arrange vbTileHorizontal
    Case 1
      Me.Arrange vbVertical
    Case 2
      Me.Arrange vbCascade
  End Select
End Sub

Private Sub munCariRegister_Click()
  Load frmCariRegisterNasabah
  frmCariRegisterNasabah.Show
End Sub

Private Sub rptRekapitulasiTabunganKeseluruhanPDL_Click()
  Load rptRekapitulasiKeseluruhanTabunganPDL
  rptRekapitulasiKeseluruhanTabunganPDL.Show
End Sub

Private Sub test_Click()
'  Load rptTest
'  rptTest.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.key
    Case "Close"
      mnuExit_Click
    Case "Printer"
      MnuCfgSetPrinter_Click
    Case "About"
      MnuAboutHelp_Click
    Case "LogOff"
      mnuLogOff_Click
    Case "Calculator"
      Shell "Calc"
    Case "Help"
      'objMenuEx.MenuDesigner (Me.hWnd)
      'MsgBox "Sorry.. Not Supported at this time!!" & vbCrLf & "Please contact this software vendor as describe in About", vbInformation
  End Select
End Sub

Private Sub GetWallpaper()
  On Error GoTo EmptyPicture:
  CommonDialog1.filter = "Picture (*.BMP;*.JPG;*.GIF) |*.BMP;*.JPG;*.GIF|"
  CommonDialog1.FileName = GetRegistry(reg_Wallpaper)
  CommonDialog1.Action = 1
  If Trim(CommonDialog1.FileName) <> "" And Dir(CommonDialog1.FileName) <> "" Then
    Me.Picture = LoadPicture(GetPicture(CommonDialog1.FileName))
    Me.Hide
    Me.Show
  End If

  SaveRegistry reg_Wallpaper, CommonDialog1.FileName
  Exit Sub
  
EmptyPicture:
  CommonDialog1.FileName = ""
  Me.Picture = LoadPicture("")
  Resume Next
End Sub

Private Function CekTrialVersionAktif() As Boolean
Dim sSql As String
Const MaxRecord  As Double = 100
CekTrialVersionAktif = True

sSql = "Select Count(kode) as JumlahRecord from registernasabah"
Set dbData = objData.SQL(GetDSN, sSql)
If dbData.RecordCount > 0 Then
  If dbData!JumlahRecord >= 50 Then
    CekTrialVersionAktif = False
    Exit Function
  End If
End If

sSql = "Select Count(Faktur) as JumlahRecord from mutasitabungan"
Set dbData = objData.SQL(GetDSN, sSql)
If dbData.RecordCount > 0 Then
  If dbData!JumlahRecord >= MaxRecord Then
    CekTrialVersionAktif = False
    Exit Function
  End If
End If

sSql = "Select Count(Faktur) as JumlahRecord from mutasideposito"
Set dbData = objData.SQL(GetDSN, sSql)
If dbData.RecordCount > 0 Then
  If dbData!JumlahRecord >= MaxRecord Then
    CekTrialVersionAktif = False
    Exit Function
  End If
End If

sSql = "Select Count(Faktur) as JumlahRecord from angsuran"
Set dbData = objData.SQL(GetDSN, sSql)
If dbData.RecordCount > 0 Then
  If dbData!JumlahRecord >= MaxRecord Then
    CekTrialVersionAktif = False
    Exit Function
  End If
End If
End Function
