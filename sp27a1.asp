<html>
<head>
<title>Laporan Arahan Kerja-Senarai Pekerja</title>
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


<form action="sp27a1.asp" method="POST">
<%Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"
  
  
id = Request.querystring("id_ak")
blnthn=request.querystring("blnthn")
tkh_ot=request.querystring("tkh_ot")
tkh_ot2=request.querystring("tkh_ot2")
proses = request.querystring("proses")
id_ak =  request.querystring ("id_ak")
ptj=plokasi

	
  ss = " select a.id_ak , a.no_pekerja , to_char(a.masa_tamat_angg,'HH24:MI') as masa_tamat , b.ptj , b.keterangan , b.pemohon , "
  ss = ss & " to_char(a.tkh_ot,'dd/mm/yyyy') as tkh_ot ,  to_char(a.masa_mula_angg,'HH24:MI') as masa_mula , to_char(a.tkh_ot,'mm/yyyy') as tkh_ot2  "
  ss = ss & " from payroll.pekerja_ot a , payroll.arahan_kerja b where a.id_ak = b.id_ak"
  ss = ss & "  and to_char(a.tkh_ot,'mm/yyyy') ='"& tkh_ot &"' "
  ss = ss & " and a.id_ak = '"&id_ak&"' order by tkh_ot , masa_mula ,masa_tamat asc  "
  
  Set rsss = objConn.Execute (ss)
 ' response.write ss
    id = rsss ("id_ak")
	tkh_ot = rsss ("tkh_ot")
	tkh_ot2 = rsss ("tkh_ot2")
	pemohon = rsss ("no_pekerja")
	pemohon2 = rsss ("pemohon")
	mula = rsss ("masa_mula")
	tamat = rsss ("masa_tamat")
	ptj = rsss ("ptj")
	keterangan = rsss ("keterangan")
	
	
	pp = "select initcap(nama) nama from payroll.paymas where no_pekerja='"&pemohon2&"' "
	Set rspp = objConn.Execute (pp)
	nama = rspp ("nama")
	
     m = " select kod , keterangan from iabs.jabatan "
     m = m & " where kod = '"&ptj&"' "
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
      LAPORAN SENARAI PEKERJA MENGIKUT ARAHAN KERJA<br> 
      JABATAN <%=ket%><br>
      BULAN <%=tkh_ot2%></font></strong></td>
  <td rowspan="3" width="10%"></td>
</tr>


<tr>
  <td colspan=3>&nbsp;</td>
</tr>
</table>
<br><br>


<table align="center" border="0" cellpadding="0" cellspacing="0" width="88%">
    <tr class="hd"> 
    <td width="26%" height="20" align="left"><b>Jabatan</b></td>
    <td width="74%" align="left"> : <%=ptj%> - <%=ket%></td>
  </tr>
    <tr class="hd"> 
    <td width="26%" align="left"><b>No Rujukan Arahan Kerja</b></td>
    <td width="74%" align="left"> : <%=id_ak%></td>
  </tr>
      <tr class="hd"> 
    <td width="26%" align="left"><b>Nama Pemohon</b></td>
    <td width="74%" align="left"> : <%=nama%> (<%=pemohon2%>)</td>
</tr>
     <tr class="hd"> 
    <td width="26%" align="left"><b>Keterangan Kerja</b></td>
    <td width="74%" align="left"> : <%=keterangan%></td>
  </tr>
</table>

<br><br><br>
<table align="center" border="0" cellpadding="0" cellspacing="0" width="90%">
  <tr class="hd"> 
    <td width="3%"  align="center"><b>Bil</b></td>
    <td width="9%"  align="center"><b>Arahan Kerja</b></td>
    <td width="8%"  align="center"><b>Tarikh</b></td>
    <td width="7%"  align="center"><b>Mula</b></td>
    <td width="7%"  align="center"><b>Tamat</b></td> 
    <td width="9%"  align="center"><b>No Pekerja</b></td>
    <td width="20%" align="left">&nbsp;&nbsp;<b>Nama Pekerja</b></td>
  </tr>
</table>
<hr width="90%">
  
  
<%  
bil=0
Do While Not rsss.eof 
bil=bil + 1
tkh_ot = rsss("tkh_ot")
id = rsss ("id_ak")	
pemohon = rsss ("no_pekerja")
mula = rsss ("masa_mula")
tamat = rsss ("masa_tamat")

	
   se = " select nvl(angg_ot_S,0) + nvl(angg_ot_M,0) anggaran "
   se = se & " from payroll.pekerja_ot"
   se = se & " where id_ak='"&id&"' and no_pekerja='"&pemohon&"' and  to_char(tkh_ot,'dd/mm/yyyy')='"&tkh_ot&"' and to_char(masa_mula_angg,'HH24:MI')= '"&mula&"' " 
   Set rse= objConn.Execute (se)
   'response.write se
   anggaran = rse("anggaran")
  

  p = "  select id_ak, keterangan, nota , ptj "
  p = p & "  from payroll.arahan_kerja "
  p = p & " where id_ak = '"&id&"' "
  Set rsp = objConn.Execute(p)
   Do While Not rsp.eof 
  keterangan1 = rsp ("keterangan")
  nota = rsp ("nota")
  ptj = rsp ("ptj")
  

  n = " select initcap(nama) nama , jawatan from payroll.paymas " 
  n = n & " where no_pekerja = '"&pemohon&"' " 
  Set rsn = objConn.Execute (n)
  
   Do While Not rsn.eof 
	nama1 = rsn ("nama")
	jaw = rsn ("jawatan")
%>


<table align="center" border="0" cellpadding="0" cellspacing="0" width="90%">
  <tr bgcolor="">
    <td width="3%"  rowspan="2" align="center"><font size="3"><%=bil%></font></td>
    <td width="9%"  rowspan="2" align="center"><font size="3"><%=rsss ("id_ak")%></font></td>
    <td width="8%"  rowspan="2" align="center"><font size="3"><%=tkh_ot%></font></td>
    <td width="7%"  rowspan="2" align="center"><font size="3"><%=rsss ("masa_mula")%></font></td>
    <td width="7%"  rowspan="2" align="center"><font size="3"><%=rsss ("masa_tamat")%></font></td>
    <td width="9%"  rowspan="2" align="center"><font size="3"><%=rsss("no_pekerja")%></font></td>
    <td width="20%" rowspan="2" align="left"><font size="3">&nbsp;&nbsp;<%=nama1%></font></td>
  </tr>
</table>

<% 
rsn.movenext
loop
rsp.movenext
loop
rsss.movenext
loop 
%>

<% 
   anggaran = rse("anggaran")	
   sss = " select sum(nvl(a.angg_ot_s,0) + nvl(a.angg_ot_m,0)) semuaxgst1 "
   sss = sss & " from payroll.pekerja_ot a , payroll.arahan_kerja c "
   sss = sss & " where a.id_ak = c.id_ak and to_char(a.tkh_ot,'mm/yyyy') ='"& tkh_ot2 &"'  " 
   sss = sss & " and c.id_ak = '"&id_ak&"'   "
   Set rssss = objConn.Execute (sss)
   'response.write sss
   semuaxgst1 = rssss("semuaxgst1")
   'Do While Not rssss.eof 
   'jum_ot_semua = jum_ot_semua + cdbl(rssss("semuaxgst1"))
 
  %>
   
  
<% 
'rssss.movenext
'loop 
%> 

<hr width="90%"> 
<table align="center" border="0" cellpadding="0" cellspacing="0" width="90%">
</table>
<hr width="90%"> 



