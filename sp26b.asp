<html>
<head>
<title>Cetakan Permohonan Seorang Pekerja</title>
<STYLE TYPE="text/css">
table.page-break{page-break-after:always}
</STYLE>
</head>
<body>



<form action="sp26b.asp" method="POST">

<%Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"

   id = Request.querystring("id_ak")
  
  
  p = "  select id_ak , penyelia , pemohon, pengesahan, tkh_sah, unit, ptj, keterangan, nota, tkh_input, tahun, bulan "
  p = p & "  from payroll.arahan_kerja where id_ak='"&id&"' "
  Set rsp = objConn.Execute(p)
 'response.write p
  id = rsp ("id_ak")
  bln = rsp ("bulan")
  thn = rsp ("tahun")
  unit = rsp ("unit")
  pemohon = rsp ("pemohon")
  penyelia = rsp ("penyelia")
  keterangan = rsp ("keterangan")
  nota = rsp ("nota")
  ptj = rsp ("ptj")
  
  
  
  m = " select kod , keterangan from iabs.jabatan "
  m = m & " where kod = '"&ptj&"' "
  Set rsm = objConn.Execute(m)
  ket = rsm ("keterangan")
  


  qpp = " select no_pekerja , nama , jawatan from payroll.paymas " 
  qpp = qpp & " where no_pekerja = '"&penyelia&"' " 
  Set rsqpp = objConn.Execute (qpp)
	nama3 = rsqpp ("nama")
	jaw1 = rsqpp ("jawatan")	
	
	
	
  pp = " select no_pekerja , nama , jawatan from payroll.paymas " 
  pp = pp & " where no_pekerja = '"&pemohon&"' " 
  Set rspp = objConn.Execute (pp)
	nama1 = rspp ("nama")
	jaw1 = rspp ("jawatan")
	
	
  ss = " select distinct no_pekerja "
  ss = ss & " from payroll.pekerja_ot " 
  ss = ss & " where id_ak = '"&id&"' order by no_pekerja" 
  Set rsss = objConn.Execute (ss)
 ' response.write rsss
    Do While Not rsss.eof 
		no_pekerja = rsss ("no_pekerja")
	
	
  n = " select no_pekerja , nama , jawatan from payroll.paymas " 
  n = n & " where no_pekerja = '"&no_pekerja&"' " 
  Set rsn = objConn.Execute (n)
	nama = rsn ("nama")
	jaw = rsn ("jawatan")
%>


<table align="center" border="0" cellpadding="0" cellspacing="0" width="98%">
<tr>
  <td align="center" width="10%" valign="top" rowspan="3" >
  <img border="0" src="images/logompsp.jpg"></td>
  <td valign="top" align="center">
      <strong><font size="3" face="Verdana">
      MAJLIS PERBANDARAN SEBERANG PERAI<br>
      PERMOHONAN KERJA LEBIH MASA<br>JABATAN&nbsp;<%=ket%>
      <br>ARAHAN KERJA - <%=id%></font></strong></td>
  <td rowspan="3" width="10%"></td>
</tr>


<tr>
  <td colspan=3>&nbsp;</td>
</tr>
</table>


<br>
<table align="center" border="0" cellpadding="0" cellspacing="0" width="78%">
<tr>
  <td width="31%" ><font face="Verdana" size="2"><b>&nbsp;Arahan Kerja</b></font></td>
  <td width="2%" ><font face="Verdana" size="2"><b>&nbsp;:</b></font></td>
  <td width="67%" ><font face="Verdana" size="2"><input type="hidden" value="<%=id%>" name="id"><%=id%></font></td>
</tr>

<tr>
  <td width="31%" ><font face="Verdana" size="2"><b>&nbsp;Nama Pekerja</b></font></td>
  <td width="2%" ><font face="Verdana" size="2"><b>&nbsp;:</b></font></td>
  <td width="67%" ><font face="Verdana" size="2"><%=nama%><input type="hidden" name="nama" value="<%=nama%>"></font></td>
</tr>

<tr>
  <td ><font face="Verdana" size="2"><b>&nbsp;No Pekerja</b></font></td>
  <td ><font face="Verdana" size="2"><b>&nbsp;:</b></font></td>
  <td ><font face="Verdana" size="2"><%=no_pekerja%></font></td>
</tr>

<tr><td colspan="3">&nbsp;</td></tr>
<tr><td colspan="3">&nbsp;</td></tr>
<tr><td colspan="3">&nbsp;</td></tr>
<tr><td colspan="3">&nbsp;</td></tr>
<tr><td colspan="3">&nbsp;</td></tr>
<tr>
  <td colspan="3">Dengan ini tuan/puan diarah menjalankan kerja lebih masa seperti berikut :</td></tr>
</table>


<table align="center" border="0" cellpadding="0" cellspacing="0" width="98%">
<tr>
  <td align="center" valign="top" colspan="3">
  </td>
</tr>
<tr><td colspan="3"></td></tr>
</table>
<br>


<table align="center" border="1" cellpadding="0" cellspacing="0" width="80%">
  <tr class="hd"> 
    <td align="center" bgcolor="#CCCCCC"><b>Tarikh</b></td>
    <td align="center" bgcolor="#CCCCCC"><b>Masa Mula (HH24:MM)</b></td>
    <td align="center" bgcolor="#CCCCCC"><b>Masa Tamat (HH24:MM)</b></td>
    <td align="center" bgcolor="#CCCCCC"><b>Keterangan Kerja Lebih Masa</b></td>
  </tr>


<%
  id = rsp ("id_ak")
  no_pekerja = rsss ("no_pekerja")
  anggaran = 0
  
  qq = " select to_char(masa_mula_angg,'HH24:MI') as masa_mula , to_char(masa_tamat_angg,'HH24:MI') as masa_tamat ,to_char(tkh_ot,'dd/mm/yyyy') as tkh_ot , "
  qq = qq & " no_pekerja , angg_ot_s , angg_ot_m  "
  qq = qq & " from payroll.pekerja_ot " 
  qq = qq & " where id_ak = '"&id&"' and no_pekerja ='"&no_pekerja&"' "
  qq =qq & " order by tkh_ot, masa_mula , masa_tamat asc  " 
  Set rsqq = objConn.Execute (qq)
   'response.write qq
   Do While Not rsqq.eof 
	mula = rsqq ("masa_mula")
	tamat = rsqq ("masa_tamat")
	tkh_ot = rsqq ("tkh_ot")
	ot_s = rsqq ("angg_ot_s")
	ot_m = rsqq ("angg_ot_m")

	anggaran = cdbl(ot_s) +  cdbl(ot_m)

%>

   <tr>
   <td align="center">&nbsp;<%=tkh_ot%></td>
   <td align="center">&nbsp;<%=mula%></td>
   <td align="center">&nbsp;<%=tamat%></td>
   <td align="center">&nbsp;<%=keterangan%></td>
   </tr> 
   <%  
    rsqq.movenext
	loop 
	 ' rsss.movenext
	'loop 
   %>
</table>           
<br><br>


<table align="center" border="0" cellpadding="0" cellspacing="0" width="85%">
<tr>
  <td width="34%" ><font face="Verdana" size="2"><b>&nbsp;Tandatangan Pekerja</b></font></td>
  <td width="2%" ><font face="Verdana" size="2"><b>&nbsp;:</b></font></td>
  <td width="64%" height="70"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br><br><br>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
  <%=nama%>-(<%=no_pekerja%>)</font></td>
</tr>
<br>

<tr>
  <td ><font face="Verdana" size="2"><b>Pegawai Yang <br>
  Meluluskan</b></font></td>
  <td ><font face="Verdana" size="2"><b>&nbsp;:</b></font></td>
  <td height="70" ><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <br><br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
  <%=nama3%>-(<%=penyelia%>)</font></td>
</tr>

<tr><td colspan="3">&nbsp;</td></tr>
<tr><td colspan="3">&nbsp;</td></tr>
<tr><td colspan="3">&nbsp;</td></tr>
<tr><td colspan="3">&nbsp;</td></tr>
<tr><td colspan="3">&nbsp;</td></tr>
</table>



<table class="page-break" align="center" border="0" cellpadding="0" cellspacing="0" width="80%">
<tr>
  <td align="center" valign="top" colspan="3">
  </td>
</tr>
<tr><td colspan="3"><font face="Verdana" size="2">Rujukan Pekeliling Perbendaharaan Bil. 7 Tahun 2008</font></td></tr>
<tr>
  <td colspan="5"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"Semua arahan kerja lebih masa hendaklah dikeluarkan secara bertulis sebelum sesuatu tempoh kerja lebih masa dimulakan"
 </font></td></tr>
<tr><td colspan="3">&nbsp;</td></tr>
<tr><td colspan="3">&nbsp;</td></tr>
<tr><td colspan="3">&nbsp;</td></tr>
<%  
rsss.movenext
loop 
%>
</table>





