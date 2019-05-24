<html>
<head>
<title>Laporan Arahan Kerja-Semakan Pengesahan Mengikut Bulan & Pemohon</title>
</head>
<body>
<script type="text/javascript" src="menu1.js"></script>
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
function check(f){
	if (f.blnthn.value=="")
	{
		alert("Sila masukkan bulan dan tahun arahan kerja");
		f.blnthn.focus();
		return false
	}
}
//  End -->
</script>



<form action="sp27b1.asp" method="POST">
<%Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"


id = Request.querystring("id_ak")
blnthn=request.querystring("blnthn")
tkh_ot=request.querystring("tkh_ot")
pemohon = request.querystring("pemohon")
proses = request.querystring("proses")
ptj = plokasi
 
 
    'check jabatan by id login
    pekd = request.cookies("gnop")
    a =     "select jabatan from majlis.pengguna where no_pekerja = '"& pekd &"' "
    Set rsa = objConn.Execute(a)

   'details by id login
   idd = "select a.nama, a.lokasi, b.keterangan, to_char(sysdate,'dd/mm/yyyy') tkh  from payroll.paymas a, iabs.jabatan b "
   idd = idd & " where a.no_pekerja = '"&pekd&"' and b.kod = a.lokasi "
   set idd2 = objConn.Execute(idd)
   'response.write idd
   lokasi = idd2 ("lokasi")
 
   'papar data pengesahan ---> nadia (16032016)
   st = " select distinct b.id_ak , to_char(b.tkh_ot,'mm/yyyy') as tkh_ot , a.ptj "
   st = st & " from payroll.jadual_ot b , payroll.arahan_kerja a  where a.id_ak=b.id_ak and  to_char(b.tkh_ot,'mm/yyyy') ='"& tkh_ot&"' and a.ptj= '"&lokasi&"' "
   st = st & " and a.pemohon='"&pemohon&"' order by id_ak"
   'response.write st
   Set rst = objConn.Execute (st)
   tkh_ot = rst ("tkh_ot")
   id_ak = rst ("id_ak")

   'keterangan jabatan    
   m = " select kod , keterangan from iabs.jabatan "
   m = m & " where kod = '"&lokasi&"' "
   Set rsm = objConn.Execute(m)
   ket = rsm ("keterangan")	
  
%>


<table align="center" border="0" cellpadding="0" cellspacing="0" width="98%">
<tr>
  <td align="center" width="10%" valign="top" rowspan="3" >
  <img border="0" src="images/logompsp.jpg"></td>
  <td valign="top" align="center">
      <strong><font size="3" face="Verdana">
      MAJLIS PERBANDARAN SEBERANG PERAI<br>
      LAPORAN PENDAFTARAN ARAHAN KERJA<br>
      JABATAN <%=ket%><br>
       BULAN <%=tkh_ot%>
     </font></strong></td>
  <td rowspan="3" width="10%"></td>
</tr>


<tr>
  <td colspan=3>&nbsp;</td>
</tr>
</table>
<br><br>
<table align="center" border="0" cellpadding="0" cellspacing="0" width="98%">
  <tr class="hd"> 
    <td width="2%"    align="center"><b>Bil</b></td>
    <td width="10%"   align="center"><b>Arahan Kerja</b></td>
    <td width="23%"   align="left"><b>Nama Pemohon</b></td>
    <td width="9%"    align="center"><b>Tarikh Sah</b></td>
    <td width="9%"    align="center"><b>Jabatan</b></td>
    <td width="8%"    align="center"><b>Pengesahan</b></td>
    <td width="10%"   align="center"><b>Jumlah Pekerja</b></td>
    <td width="13%"   align="center"><b>Anggaran OT (RM)</b></td>
    <td width="16%"   align="center"><b>Tuntutan OT Sebenar (RM)</b></td>
</table>
<hr width="98%">
  
  
  
<%  
bil=0
Do While Not rst.eof 
id_ak = rst ("id_ak")
tkh_ot = rst ("tkh_ot")
bil=bil+ 1



		ss = " select  id_ak , ptj , to_char(tkh_sah,'dd/mm/yyyy') as tkh_sah , pengesahan  , bulan , tahun , pemohon "
		ss = ss & " from payroll.arahan_kerja where "
		ss =ss & "  ptj='"&lokasi&"' and id_ak='"&id_ak&"' "
		ss = ss & " order by id_ak , ptj " 
		'response.write ss
		Set rsss = objConn.Execute (ss)
		id_ak = rsss ("id_ak")
		ptj = rsss ("ptj")
		pengesahan = rsss ("pengesahan")
		tkh_sah = rsss ("tkh_sah")
		pemohon = rsss ("pemohon")



   		pp = " select no_pekerja , initcap(nama) nama , jawatan from payroll.paymas " 
   		pp = pp & " where no_pekerja = '"&pemohon&"' " 
  		Set rspp = objConn.Execute (pp)
   		nama1 = rspp ("nama")
   		jaw1 = rspp ("jawatan")	
   
	
 		m = " select kod , initcap(keterangan) keterangan from iabs.jabatan "
  		m = m & " where kod = '"&ptj&"' "
  		Set rsm = objConn.Execute(m)
  		ket = rsm ("keterangan")	
  
  

   		se = " select sum(nvl(angg_ot_s,0) + nvl(angg_ot_m,0)) anggaran , count (distinct no_pekerja) as kira , sum(nvl(ot_sebenar_m,0) + nvl(ot_sebenar_s,0)) sebenar"
   		se = se & " from payroll.pekerja_ot"
   		se = se & " where id_ak='"&id_ak&"'  " 
   		Set rse= objConn.Execute (se)
   		'response.write se
     	Do While Not rse.eof 
   		anggaran = rse("anggaran")
   		kira = rse("kira")
   		sebenar = rse("sebenar")
%>



<table align="center" border="0" cellpadding="0" cellspacing="0" width="98%">
  <tr class="hd"> 
    <td width="2%"   rowspan="2"  align="center"><%=bil%></td>
    <td width="10%"  rowspan="2"  align="center"><%=rst ("id_ak")%></td>
    <td width="23%"  rowspan="2"  align="left"><%=rsss ("pemohon")%>- <%=rspp ("nama")%></td>
    <td width="9%"   rowspan="2"  align="center"><%=rsss ("tkh_sah")%></td>
    <td width="9%"   rowspan="2"  align="center"><%=rsss ("ptj")%></td>
    <td rowspan="2" align="center"> 
	<%if pengesahan="Y" then%>Y<%end if%>
	<%if pengesahan="T" then%><b>T</b><%end if%>
    </td>
    <td width="10%" rowspan="2"  align="center"><%=rse ("kira")%></td>
    <td width="13%" rowspan="2"  align="center"><%=rse("anggaran")%></td>
    <td width="16%" rowspan="2"  align="center"><%=rse("sebenar")%></td>
  </tr>
</table>


<%  
rse.movenext
loop 
rst.movenext
loop
%>
 
   <%    
   sss = " select sum(nvl(a.angg_ot_s,0) + nvl(a.angg_ot_m,0)) semuaxgst1 , sum(nvl(ot_sebenar_m,0) +  nvl(ot_sebenar_s,0)) sebenar1 "
   sss = sss & " from payroll.pekerja_ot a , payroll.arahan_kerja c "
   sss = sss & " where a.id_ak = c.id_ak and  to_char(a.tkh_ot,'mm/yyyy') ='"& tkh_ot&"' and c.ptj='"&ptj&"' and c.pemohon='"&pemohon&"' "  
   Set rssss = objConn.Execute (sss)
    'response.write sss
   semuaxgst1 = rssss("semuaxgst1")
   sebenar1 = rssss("sebenar1")
  %> 
  
<hr width="98%"> 
<table align="center" border="0" cellpadding="0" cellspacing="0" width="98%">
  <tr bgcolor="">
    <td colspan="5" align="right"><B>&nbsp; Jumlah Keseluruhan Anggaran OT (RM)&nbsp;</B></td>
    <td width="13%" align="center"><b><%=semuaxgst1%></b></td>
    <td width="16%" align="center"><b><%=sebenar1%></b></td>
  </tr>
</table>
<hr width="98%"> 



