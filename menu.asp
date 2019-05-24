<!--#include file="connection.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
	"http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- using doc write to read query string
<link type="text/css" rel="StyleSheet" href="skins/officexp/officexp.css" />

-->

<script type="text/javascript">


var ie55 = /MSIE ((5\.[56789])|([6789]))/.test( navigator.userAgent ) &&
			navigator.platform == "Win32";

if ( !ie55 ) {
	window.onerror = function () {
		return true;
	};
}

function writeNotSupported() {
	if ( !ie55 ) {
		document.write( "<p class=\"warning\">" +
			"This script only works in Internet Explorer 5.5" +
			" or greater for Windows</p>" );
	}
}

</script>
<script type="text/javascript">

function getQueryString( sProp ) {
	var re = new RegExp( sProp + "=([^\\&]*)", "i" );
	var a = re.exec( document.location.search );
	if ( a == null )
		return "";
	return a[1];
};

function changeCssFile( sCssFile ) {
	var loc = String(document.location);
	var search = document.location.search;
	if ( search != "" )
		loc = loc.replace( search, "" );
	loc = loc + "?css=" + sCssFile;
	document.location.replace( loc );
}

var cssFile = getQueryString( "css" );
if ( cssFile == "" )
	//cssFile = "skins/winclassic.css";
	cssFile = "skins/winclassic.css";
cssFile = "skins/officexp/officexp.css";
document.write("<link type=\"text/css\" rel=\"StyleSheet\" href=\"" + cssFile + "\" />" );

</script>
<script language="javascript">
function pasti(){
var pasti=confirm("Pasti Log Out?");
if (pasti==true) {
document.location.href ="main.asp"
return true;
}
else return false;
}

function pasti1(){
var pasti=confirm("Pasti Keluar?");
if (pasti==true) {
document.location.href ="../../payroll/mpspnet/sistem.asp"
return true;
}
else return false;
}
</script>
<style type="text/css">


.menu-bar {
	border-bottom:	2px groove;
	
}

p {
	font:	Message-Box;
	font:	MessageBox;
	margin:	10px;
}

.warning {
	color:	red;
}

a {
	color: blue;
}

textarea {
	margin:		10px;
	display:	block;
	width:		auto;
	
}

</style>
<%

   
   
response.expires=-1
idstaff = request.cookies("gnop")
pwd = request.cookies("pass")
mula = request.cookies("mula")
kod_sistem="k"

Session("np") = idstaff
Session("pass") = pwd
Session("mula") = mula

A = Request.ServerVariables("SCRIPT_NAME")
pjg=len(A)
for i = 1 to pjg
temp=mid(A,i,1)
	if temp="/" then MyPos2=i
next
page=mid(A,MyPos2 + 1,(pjg-Mypos2-4))

if idstaff <> "" and pwd <> "" and kod_sistem <> "" then

  if page <> "index" and page <> "main" and page <> "komen1" and page <> "report2" and page <> "report3" and page <> "komen2" and page <> "peti_surat_ind" then
  c = "select 'x' from kebenaran_2002 a,menu_kehadiran b where a.no_pekerja = "& idstaff &" and "&_
  "a.sistem like '"& kod_sistem &"' and a.skrin like b.kod and b.perkara like '"& page &"'"
  Set v = objConn1.Execute(c)
  	if v.eof then 
	session.abandon
	response.cookies("gnop").expires = Date()-1
	response.cookies("pass").expires = Date()-1
		if mula="index" then
		response.redirect "http://mpspnet.mpsp.gov.my/sistemnet.asp"
		else
		response.redirect "main.asp"
		end if
	end if
  end if
  

c1="select skrin from majlis.kebenaran_2002 where no_pekerja = "& idstaff &" and "&_
"sistem like '"& kod_sistem &"'"
set v1 = objConn.Execute(c1)

if not v1.eof then
	do while not v1.eof
	kodabc = v1("skrin")
	
	if mid(kodabc,1,3) ="k01" then
		k01 = "Laporan"
		if kodabc = "k0101" then kodk0101 = kodabc
		if kodabc = "k0102" then kodk0102 = kodabc
		if kodabc = "k0103" then kodk0103 = kodabc
		if kodabc = "k0104" then kodk0104 = kodabc
		if kodabc = "k0105" then kodk0105 = kodabc
		if kodabc = "k0106" then kodk0106 = kodabc
		if kodabc = "k0107" then kodk0107 = kodabc
		if kodabc = "k0108" then kodk0108 = kodabc
		if kodabc = "k0109" then kodk0109 = kodabc
		if kodabc = "k0110" then kodk0110 = kodabc
		if kodabc = "k0111" then kodk0111 = kodabc
		if kodabc = "k0112" then kodk0112 = kodabc
		if kodabc = "k0113" then kodk0113 = kodabc
		if kodabc = "k0114" then kodk0114 = kodabc
		if kodabc = "k0115" then kodk0115 = kodabc
		if kodabc = "k0116" then kodk0116 = kodabc

	end if
	
	if mid(kodabc,1,3) ="k02" then
		k02 = "Komen"
		if kodabc = "k0201" then kodk0201 = kodabc
		if kodabc = "k0202" then kodk0202 = kodabc
		if kodabc = "k0203" then kodk0203 = kodabc
		if kodabc = "k0204" then kodk0204 = kodabc
		if kodabc = "k0205" then kodk0205 = kodabc
	end if
	
	if mid(kodabc,1,3) ="k03" then
		k03 = "Daftar"
		if kodabc = "k0301" then kodk0301 = kodabc
		if kodabc = "k0302" then kodk0302 = kodabc
		if kodabc = "k0303" then kodk0303 = kodabc
		if kodabc = "k0304" then kodk0304 = kodabc
		if kodabc = "k0305" then kodk0305 = kodabc
		if kodabc = "k0306" then kodk0306 = kodabc
		if kodabc = "k0307" then kodk0307 = kodabc
	end if

	if mid(kodabc,1,3) ="k04" then
		k04 = "Selenggara"
		if kodabc = "k0401" then kodk0401 = kodabc
		if kodabc = "k0402" then kodk0402 = kodabc
		if kodabc = "k0403" then kodk0403 = kodabc
		if kodabc = "k0404" then kodk0404 = kodabc
		if kodabc = "k0405" then kodk0405 = kodabc
		if kodabc = "k0406" then kodk0406 = kodabc
		if kodabc = "k0407" then kodk0407 = kodabc
	end if
	
	if mid(kodabc,1,3) ="k05" then
		k05 = "WBB"
		if kodabc = "k0501" then kodk0501 = kodabc
		if kodabc = "k0502" then kodk0502 = kodabc
		if kodabc = "k0503" then kodk0503 = kodabc
		if kodabc = "k0504" then kodk0504 = kodabc
		if kodabc = "k0505" then kodk0505 = kodabc
		if kodabc = "k0506" then kodk0506 = kodabc
		if kodabc = "k0507" then kodk0507 = kodabc
		if kodabc = "k0508" then kodk0508 = kodabc
		if kodabc = "k0509" then kodk0509 = kodabc
		if kodabc = "k0510" then kodk0510 = kodabc
		if kodabc = "k0511" then kodk0511 = kodabc
		if kodabc = "k0512" then kodk0512 = kodabc
		if kodabc = "k0513" then kodk0513 = kodabc
	end if
	
	if mid(kodabc,1,3) ="k06" then
		k06 = "Status"
		if kodabc = "k0601" then kodk0601 = kodabc
		if kodabc = "k0602" then kodk0602 = kodabc
		if mid(kodabc,1,5) = "k0603" then		
			kodk0603 = "Statistik Kesalahan (Admin)"						
			if kodabc = "k060301" then kodk060301 = kodabc
			if kodabc = "k060302" then kodk060302 = kodabc
			if kodabc = "k060303" then kodk060303 = kodabc
		end if
		if kodabc = "k0604" then kodk0604 = kodabc
		if kodabc = "k0605" then kodk0605 = kodabc
		if kodabc = "k0606" then kodk0606 = kodabc
		if kodabc = "k0607" then kodk0607 = kodabc
		if kodabc = "k0608" then kodk0608 = kodabc
		if mid(kodabc,1,5) = "k0609" then
			kodk0609 = "Statistik Kesalahan"				
			if kodabc = "k060901" then kodk060901 = kodabc
			if kodabc = "k060902" then kodk060902 = kodabc
			if kodabc = "k060903" then kodk060903 = kodabc
		end if
		if kodabc = "k0610" then kodk0610 = kodabc
	end if
	
	if mid(kodabc,1,3) ="k07" then
		k07 = "Pengesahan"
		if kodabc = "k0701" then kodk0701 = kodabc
		if kodabc = "k0702" then kodk0702 = kodabc
		if kodabc = "k0703" then kodk0703 = kodabc
		if kodabc = "k0704" then kodk0704 = kodabc
		if kodabc = "k0705" then kodk0705 = kodabc
		if kodabc = "k0706" then kodk0706 = kodabc
		if kodabc = "k0707" then kodk0707 = kodabc
		if kodabc = "k0708" then kodk0708 = kodabc
		if kodabc = "k0709" then kodk0709 = kodabc
	end if
	
	if mid(kodabc,1,3) ="k08" then
		k08 = "eSurat"
		if kodabc = "k0801" then kodk0801 = kodabc
		if kodabc = "k0802" then kodk0802 = kodabc
		if kodabc = "k0803" then kodk0803 = kodabc
		if kodabc = "k0804" then kodk0804 = kodabc
		if kodabc = "k0805" then kodk0805 = kodabc
		if kodabc = "k0806" then kodk0806 = kodabc
		if kodabc = "k0807" then kodk0807 = kodabc
		if kodabc = "k0808" then kodk0808 = kodabc
	end if
	
	if mid(kodabc,1,3) ="k09" then
		k09 = "Cetak Surat"
		if kodabc = "k0901" then kodk0901 = kodabc
		if kodabc = "k0902" then kodk0902 = kodabc
		if kodabc = "k0903" then kodk0903 = kodabc
	end if	
	
	v1.movenext
	loop
v1.close
end if				
%>
<script type="text/javascript" src="js/p
oslib.js"></script>
<script type="text/javascript" src="js/scrollbutton.js"></script>
<script type="text/javascript" src="js/menu4.js"></script>
<script type="text/javascript">
//<![CDATA[

Menu.prototype.cssFile = cssFile;
Menu.prototype.mouseHoverDisabled = false;

var tmp;
var mb = new MenuBar;

//******************************************************************************************************
// Laporan

	var Laporan = new Menu();
	//SistemMenu.add( tmp = new MenuItem( "Bantuan Sistem" ,"#") );
	<% if kodk0101 <> "" then %>
		Laporan.add( tmp = new MenuItem( "Laporan Mengikut Jabatan (Admin)" ,"report.asp") );
	<% end if%>
	<% if kodk0107 <> "" then %>
		Laporan.add( tmp = new MenuItem( "Laporan Mengikut Jabatan" ,"report7.asp") );
	<% end if%>
	<% if kodk0114 <> "" then %>
		Laporan.add( tmp = new MenuItem( "Laporan Mengikut Lokasi (Admin)" ,"report_lokasi_admin.asp") );
	<% end if%>
	<% if kodk0113 <> "" then %>
		Laporan.add( tmp = new MenuItem( "Laporan Mengikut Lokasi" ,"report_lokasi.asp") );
	<% end if%>
	Laporan.add( tmp = new MenuItem( "Laporan Bulanan Individu" ,"report2.asp") );
	<% if kodk0109 <> "" then %>	
		Laporan.add( tmp = new MenuItem( "Laporan Bulanan Seorang Pekerja" ,"report9.asp") );
	<% end if %>	
	<% if kodk0110 <> "" then %>	
		Laporan.add( tmp = new MenuItem( "Laporan Bulanan Pekerja Tidak Aktif" ,"report10.asp") );
	<% end if %> 
	<% if kodk0103 <> "" then %>	
		Laporan.add( tmp = new MenuItem( "Laporan Kedatangan Lewat" ,"report3.asp") );
	<% end if %>	
	<% if kodk0105 <> "" then %>	
		Laporan.add( tmp = new MenuItem( "Laporan Balik Awal" ,"report5.asp") );
	<% end if %>
	<% if kodk0104 <> "" then %>	
		Laporan.add( tmp = new MenuItem( "Statistik Kedatangan Lewat Mengikut Jabatan" ,"report4.asp") );	
	<% end if %>	 
	<% if kodk0108 <> "" then %>	
		Laporan.add( tmp = new MenuItem( "Maklumat Kehadiran" ,"report8.asp") );
	<% end if %>
	<% if kodk0111 <> "" then %>	
		Laporan.add( tmp = new MenuItem( "Laporan Pekerja Dimaafkan Dan Tidak Dimaafkan" ,"report20.asp") );
	<% end if %>
	<% if kodk0112 <> "" then %>	
		Laporan.add( tmp = new MenuItem( "Laporan Mengikut Gred" ,"report21.asp") );
	<% end if %>
	<% if kodk0106 <> "" then %>	
		Laporan.add( tmp = new MenuItem( "Statistik Kehadiran Lewat Pegawai" ,"report6.asp") );
	<% end if %>
	<% if kodk0115 <> "" then %>
		Laporan.add( tmp = new MenuItem( "Laporan Ketidakhadiran (Admin)" ,"ketidakhadiran2n.asp") );
	<% end if%>
	<% if kodk0116 <> "" then %>
		Laporan.add( tmp = new MenuItem( "Laporan Ketidakhadiran Jabatan" ,"ketidakhadiran_jabatan.asp") );
	<% end if%>
	mb.add( tmp = new MenuButton( "Laporan",Laporan ) );

//***************************************************************************************************************
// Komen
	var Komen = new Menu();
	//SistemMenu.add( tmp = new MenuItem( "Bantuan Sistem" ,"#") );
	Komen.add( tmp = new MenuItem( "Catitan Alasan" ,"komen1.asp") );
	 <% if kodk0202 <> "" then %>
		Komen.add( new MenuSeparator() );
	<% end if %>
	<% if kodk0202 <> "" then %>
		Komen.add( tmp = new MenuItem( "Catitan Alasan Jabatan" ,"kerjaluar.asp") );
	<% end if %>
	<% if kodk0203 <> "" then %>
		Komen.add( tmp = new MenuItem( "Batal Catitan" ,"batal.asp") );
	<% end if %>
	<% if kodk0203 <> "" and kodk0204 <> "" then %>
		Komen.add( new MenuSeparator() );
	<% end if %>
	<% if kodk0204 <> "" then %>
		Komen.add( tmp = new MenuItem( "Catitan Alasan Admin" ,"kerjaluar_admin.asp") );
	<% end if %>
	<% if kodk0205 <> "" then %>
		Komen.add( tmp = new MenuItem( "Penyelenggaraan Jenis Alasan" ,"jenis_alasan.asp") );
	<% end if %>
	mb.add( tmp = new MenuButton( "Catitan",Komen ) );

//**********************************************************************************
<% if k03 <> "" then %>	
// Daftar
	var Daftar = new Menu();
	//SistemMenu.add( tmp = new MenuItem( "Bantuan Sistem" ,"#") );
	<% if kodk0301 <> "" then %>	
		Daftar.add( tmp = new MenuItem( "Daftar/Edit" ,"daftar.asp") );
	<% end if %>
	<% if kodk0302 <> "" then %>	
		Daftar.add( tmp = new MenuItem( "Kemaskini Info Kakitangan" ,"edit.asp") );
	<% end if %>
	<% if kodk0303 <> "" then %>	
		Daftar.add( tmp = new MenuItem( "Edit Status" ,"status.asp") );
	<% end if %>
	<% if kodk0304 <> "" then %>	
		Daftar.add( tmp = new MenuItem( "Daftar/Edit Jabatan" ,"daftarJabatan.asp") );
	<% end if %>
	<% if kodk0305 <> "" then %>	
		Daftar.add( tmp = new MenuItem( "Daftar Lokasi" ,"daftarlokasi.asp") );
	<% end if %>
	<% if kodk0306 <> "" then %>	
		Daftar.add( tmp = new MenuItem( "Daftar/Edit Kakitangan Admin" ,"daftarkakitangan_admin.asp") );
	<% end if %>	
	<% if kodk0307 <> "" then %>	
		Daftar.add( tmp = new MenuItem( "Daftar/Edit Kakitangan" ,"daftarkakitangan.asp") );
	<% end if %>	
	mb.add( tmp = new MenuButton( "Daftar",Daftar ) );
<% end if %>

//**************************************************************************************
<% if k05 <> "" then %>	
// WBB
	var WBB = new Menu();
	//SistemMenu.add( tmp = new MenuItem( "Bantuan Sistem" ,"#") );
	<% if kodk0501 <> "" then %>	
		WBB.add( tmp = new MenuItem( "Permohonan Baru","borangpemilihan.asp") );
	<% end if %>
	<% if kodk0502 <> "" then %>	
		WBB.add( tmp = new MenuItem( "Permohonan Semula","borangpemilihans.asp") );
	<% end if %>	
	<% if kodk0503 <> "" then %>	
		WBB.add( tmp = new MenuItem( "Semakan Waktu Bekerja","semak_kelulusan.asp") );
	<% end if %>
	<% if kodk0504 <> "" then %>	
		WBB.add( tmp = new MenuItem( "Permohonan Baru (Admin)","borangpemilihan_admin.asp") );
	<% end if %>
	<% if kodk0505 <> "" then %>	
		WBB.add( tmp = new MenuItem( "Permohonan Semula (Admin)","borangpemilihans_admin.asp") );
	<% end if %>
	<% if kodk0513 <> "" then %>	
		WBB.add( tmp = new MenuItem( "Permohonan Semula (Admin IT)","borangpemilihans_adminIT.asp") );
	<% end if %>
	<% if kodk0512 <> "" then %>	
		WBB.add( tmp = new MenuItem( "Semakan Waktu Bekerja (Admin)","semak_kelulusan_admin.asp") );
	<% end if %>	
	<% if kodk0506 <> "" then %>	
		WBB.add( tmp = new MenuItem( "Pengesahan Baru","senarai_utksah.asp") );
	<% end if %>
	<% if kodk0507 <> "" then %>	
		WBB.add( tmp = new MenuItem( "Pengesahan Semula","senarai_utksahs.asp") );
	<% end if %>
	<% if kodk0508 <> "" then %>	
		WBB.add( tmp = new MenuItem( "Laporan Sebelum Pengesahan","report_sblmpengesahan.asp") );
	<% end if %>
	<% if kodk0509 <> "" then %>	
		WBB.add( tmp = new MenuItem( "Laporan Selepas Pengesahan","report_slpspengesahan.asp") );
	<% end if %>
	<% if kodk0510 <> "" then %>	
		WBB.add( tmp = new MenuItem("Statistik Slps Pengesahan","statistik_jabatan.asp") );
	<% end if %>
	<% if kodk0511 <> "" then %>	
		WBB.add( tmp = new MenuItem("Statistik Sblm Pengesahan","statistik_jabatan_sblm.asp") );
	<% end if %>			
	mb.add( tmp = new MenuButton( "Waktu Bekerja",WBB ) );
<% end if %>
				
//*************************************************************************************************
<% if k04 <> "" then %>	
// Selenggara
	var Selenggara = new Menu();
	//SistemMenu.add( tmp = new MenuItem( "Bantuan Sistem" ,"#") );
	<% if kodk0401 <> "" then %>	
		Selenggara.add( tmp = new MenuItem( "Pendaftaran Pengguna" ,"user.asp") );
	<% end if %>
	<% if kodk0402 <> "" then %>	
		Selenggara.add( tmp = new MenuItem( "Ubah Masa" ,"ubahmasa.asp") );
	<% end if %>
	<% if kodk0405 <> "" then %>	
		Selenggara.add( tmp = new MenuItem( "Ubah Masa Keluar" ,"ubahmasaKeluar.asp") );
	<% end if %>
	<% if kodk0403 <> "" then %>	
		Selenggara.add( tmp = new MenuItem( "Pengarah" ,"pengarah.asp") );
	<%end if %>
	<% if kodk0404 <> "" then %>	
		//Selenggara.add( tmp = new MenuItem( "Hapus Rekod Kehadiran" ,"hapus.asp") );
	<% end if %>
	<% if kodk0406 <> "" then %>	
		Selenggara.add( tmp = new MenuItem( "Sessi Bekerja" ,"register_shift.asp") );
	<% end if %>
	<% if kodk0407 <> "" then %>	
		Selenggara.add( tmp = new MenuItem( "Senarai Pengguna" ,"senarai_pengguna.asp") );
	<% end if %>
	mb.add( tmp = new MenuButton( "Selenggara",Selenggara ) );
<% end if %>

//*************************************************************************************************
<% if k06 <> "" then %>	
// Status
	var Status = new Menu();
	<% if kodk0601 <> "" then %>
		Status.add( tmp = new MenuItem("Status Kehadiran (Admin)", "status_hadir_admin.asp") );
	<% end if %>
	<% if kodk0602 <> "" then %>	
		Status.add( tmp = new MenuItem("Statistik Kehadiran (Admin)", "status_semua_admin.asp") );
	<% end if %>	
	<% if kodk0603 <> "" then %>
		var SubStatusAdmin = new Menu();
		<% if kodk060301 <> "" then %>	
			SubStatusAdmin.add( tmp = new MenuItem("Statistik Lewat/Balik Awal/Tidak Scan Keluar (Admin)", "status_salah_admin.asp") );
		<% end if %>
		<% if kodk060302 <> "" then %>	
			SubStatusAdmin.add( tmp = new MenuItem("Statistik Kehadiran Lewat (Admin)", "status_lewat_admin.asp") );
		<% end if %>
		<% if kodk060303 <> "" then %>	
			SubStatusAdmin.add( tmp = new MenuItem("Statistik Balik Awal/Tidak Scan Keluar (Admin)", "status_awal_admin.asp") );
		<% end if %>
		Status.add( tmp = new MenuItem( "Statistik Kesalahan (Admin)",null,null,SubStatusAdmin ) );
	<% end if %>
	<% if kodk0604 <> "" then %>
		Status.add( tmp = new MenuItem("Statistik Tidak Hadir (Admin)", "status_xhadir_admin.asp") );
	<% end if %>
	<% if kodk0605 <> "" then %>
		Status.add( tmp = new MenuItem("Statistik Kehadiran Individu (Admin)", "status_ind_lapor.asp") );
	<% end if %>
	<% if kodk0606 <> "" then %>
		Status.add( tmp = new MenuItem("Statistik Kehadiran Keseluruhan Jabatan (Admin)", "status_hadir_lapor.asp") );
	<% end if %>
	<% if kodk0606 <> "" and kodk0607 <> "" then %>
		Status.add( new MenuSeparator() );
	<% end if %>
	<% if kodk0607 <> "" then %>
		Status.add( tmp = new MenuItem("Status Kehadiran", "status_hadir.asp") );
	<% end if %>
	<% if kodk0608 <> "" then %>	
		Status.add( tmp = new MenuItem("Statistik Kehadiran", "status_semua.asp") );
	<% end if %>
	<% if kodk0609 <> "" then %>
		var SubStatus = new Menu();
		<% if kodk060901 <> "" then %>	
			SubStatus.add( tmp = new MenuItem("Statistik Lewat/Balik Awal/Tidak Scan Keluar", "status_salah.asp") );
		<% end if %>
		<% if kodk060902 <> "" then %>	
			SubStatus.add( tmp = new MenuItem("Statistik Kehadiran Lewat", "status_lewat.asp") );
		<% end if %>
		<% if kodk060903 <> "" then %>	
			SubStatus.add( tmp = new MenuItem("Statistik Balik Awal/Tidak Scan Keluar", "status_awal.asp") );
		<% end if %>
		Status.add( tmp = new MenuItem( "Statistik Kesalahan",null,null,SubStatus ) );
	<% end if %>
	<% if kodk0610 <> "" then %>
		Status.add( tmp = new MenuItem("Statistik Tidak Hadir", "status_xhadir.asp") );
	<% end if %>
	mb.add( tmp = new MenuButton( "Status",Status ) );
<% end if %>

//*********************************************************************************************************
<% if k07 <> "" then %>	
// Sah
	var Sah = new Menu();
	<% if kodk0701 <> "" then %>
		Sah.add( tmp = new MenuItem("Pengesahan Kehadiran Lewat/Balik Awal/Tidak Scan Keluar", "sah.asp") );
	<% end if %>
	<% if kodk0709 <> "" then %>
		Sah.add( tmp = new MenuItem("Pengesahan Kehadiran Lewat Mengikut Hari", "sahByHari.asp") );
	<% end if %>
	<% if kodk0702 <> "" then %>	
		Sah.add( tmp = new MenuItem("Laporan Pengesahan Kehadiran Lewat/Balik Awal/Tidak Scan Keluar", "sah_lapor.asp") );
	<% end if %>
	<% if kodk0708 <> "" then %>	
		Sah.add( tmp = new MenuItem("Laporan Aktiviti Pengesahan", "sah_aktiviti_lapor.asp") );
	<% end if %>
	<% if kodk0707 <> "" then %>	
		Sah.add( tmp = new MenuItem("Penyerahan Sementara Tugas Pengesahan", "pass_tugas_sah.asp") );
	<% end if %>
	<% if kodk0702 <> "" and kodk0703 <> ""  then %>
		Sah.add( new MenuSeparator() );
	<% end if %>
	<% if kodk0703 <> "" then %>	
		Sah.add( tmp = new MenuItem("Pengesahan Kehadiran Lewat/Balik Awal/Tidak Scan Keluar (SUP)", "sah_sup.asp") );
	<% end if %>
	<% if kodk0704 <> "" then %>	
		Sah.add( tmp = new MenuItem("Laporan Pengesahan Kehadiran Lewat/Balik Awal/Tidak Scan Keluar (SUP)", "sah_sup_lapor.asp") );
	<% end if %>
	<% if kodk0704 <> "" and kodk0705 <> ""  then %>
		Sah.add( new MenuSeparator() );
	<% end if %>
	<%if kodk0705 <> "" then %>	
		Sah.add( tmp = new MenuItem("Laporan Pengesahan Kehadiran Lewat/Balik Awal/Tidak Scan Keluar (Admin)", "sah_admin_lapor.asp") );
	<% end if %>
	<% if kodk0706 <> "" then %>	
		Sah.add( tmp = new MenuItem("Laporan Aktiviti Pengesahan (Admin)", "sah_aktiviti_lapor_admin.asp") );
	<% end if %>
	mb.add( tmp = new MenuButton( "Pengesahan",Sah) );
<% end if %>

//********************************************************************************************************************************	
// eSurat
	var eSurat = new Menu();
	eSurat.add( tmp = new MenuItem("Peti Surat Individu", "peti_surat_ind.asp") );
	<% if kodk0802 <> "" then %>	
		eSurat.add( tmp = new MenuItem("Peti Surat Ketua Jabatan (Surat Tunjuk Sebab)", "peti_surat_peg.asp") );
	<% end if %>
	<% if kodk0803 <> "" then %>	
		eSurat.add( tmp = new MenuItem("Peti Surat Ketua Jabatan (Surat Tindakan Tatatertib)", "peti_surat_peg2.asp") );
	<% end if %>
	<% if kodk0804 <> "" then %>	
		eSurat.add( tmp = new MenuItem("Peti Surat SUP", "peti_surat_sup.asp") );
	<% end if %>
	<% if kodk0805 <> "" then %>	
		eSurat.add( tmp = new MenuItem("Peti Surat Admin", "peti_surat_admin.asp") );
	<% end if %>
	<% if kodk0805 <> "" and kodk0806 <> ""  then %>
		eSurat.add( new MenuSeparator() );
	<% end if %>
	<% if kodk0806 <> "" then %>
		eSurat.add( tmp = new MenuItem("Laporan e-Surat", "surat_lapor.asp") );
	<% end if %>
	<% if kodk0807 <> "" then %>
		eSurat.add( tmp = new MenuItem("Laporan e-Surat (SUP)", "surat_lapor_sup.asp") );
	<% end if %>
	<% if kodk0808 <> "" then %>
		eSurat.add( tmp = new MenuItem("Laporan e-Surat (Admin)", "surat_lapor_admin.asp") );
	<% end if %>
	mb.add( tmp = new MenuButton( "eSurat",eSurat) );

//******************************************************************************************
<% if k09 <> "" then %>	
// Cetak Surat
	var cetak = new Menu();
	<% if kodk0901 <> "" then %>
		cetak.add( tmp = new MenuItem("Surat Tunjuk Sebab", "surat_tunjuk_sebab.asp") );
	<% end if %>
	<% if kodk0902 <> "" then %>	
		cetak.add( tmp = new MenuItem("Surat Amaran", "surat_amaran.asp") );
	<% end if %>
	<% if kodk0903 <> "" then %>	
		cetak.add( tmp = new MenuItem("Surat Tindakan Tatatertib", "surat_tindakan.asp") );
	<% end if %>
	mb.add( tmp = new MenuButton( "Cetak Surat",cetak) );
<% end if %>

//*******************************************************************************************
// Sistem Menu
	var SistemMenu = new Menu();
	//SistemMenu.add( tmp = new MenuItem( "Bantuan Sistem" ,"#") );
	<% if mula="index" then %>
		//SistemMenu.add( tmp = new MenuItem("Status Kehadiran", "status_hadir.asp") );
		SistemMenu.add( tmp = new MenuItem( "Sistem Lain" ,"../sistem.asp") );
		SistemMenu.add( tmp = new MenuItem( "Logout Sistem" ,"http://mpspnet.mpsp.gov.my/sistemnet.asp") );
	<% else %>
		SistemMenu.add( tmp = new MenuItem( "Logout Sistem" ,"tutup.asp") );
	<% end if %>
	mb.add( tmp = new MenuButton( "Sistem",SistemMenu ) );

//*******************************************************************************************
<%Else
session.Abandon
Session.Contents.RemoveAll()
	response.cookies("gnop").expires = Date()-1
	response.cookies("pass").expires = Date()-1
		if mula="index" then
		response.redirect "http://mpspnet.mpsp.gov.my/sistemnet.asp"
		else
		response.redirect "main.asp"
		end if
End if%>

</script>