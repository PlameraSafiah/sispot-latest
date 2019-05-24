<html>
<head>
<title>Cetakan Permohonan Seorang Pekerja</title>
<style type="text/css">
table.page-break{ 
border: none;
page-break-after:always}
</style>
</head>
<body>



<form action="sp26c.asp" method="POST">

<%Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"

   id = Request.querystring("id_ak")
  
  'check id arahan kerja
  p = "  select id_ak , penyelia , pemohon, pengesahan, tkh_sah, unit, ptj, keterangan,keterangan_1, nota, tkh_input, tahun, bulan "
  p = p & "  from payroll.arahan_kerja where id_ak='"&id&"' "
  Set rsp = objConn.Execute(p)
  'response.write p
  id = rsp ("id_ak")
  bln = rsp ("bulan")
  thn = rsp ("tahun")
  pemohon = rsp ("pemohon")
  penyelia = rsp ("penyelia")
  keterangan = rsp ("keterangan")
  keterangan_1 = rsp ("keterangan_1")
  nota = rsp ("nota")
  ptj = rsp ("ptj")
  
  
  'nama jabatan
  m = " select kod , keterangan from iabs.jabatan "
  m = m & " where kod = '"&ptj&"' "
  Set rsm = objConn.Execute(m)
  ket = rsm ("keterangan")
  
	
  n = " select no_pekerja , nama , jawatan from payroll.paymas " 
  n = n & " where no_pekerja = '"&penyelia&"' " 
  Set rsn = objConn.Execute (n)
	nama3 = rsn ("nama")
	jaw3 = rsn ("jawatan")
	
	
   'nama pekerja
  pp = " select no_pekerja , nama , jawatan from payroll.paymas " 
  pp = pp & " where no_pekerja = '"&pemohon&"' " 
  Set rspp = objConn.Execute (pp)
	nama1 = rspp ("nama")
	jaw1 = rspp ("jawatan")


  ss = " select distinct no_pekerja "
  ss = ss & " from payroll.pekerja_ot " 
  ss = ss & " where id_ak = '"&id&"' order by no_pekerja" 
  Set rsss = objConn.Execute (ss) 
 ' response.write ss
   no_pekerja = rsss ("no_pekerja")	
   
   
  ssr = " select count(distinct no_pekerja) as kira"
  ssr = ssr & " from payroll.pekerja_ot " 
  ssr = ssr & " where id_ak = '"&id&"' order by no_pekerja" 
  Set rsssr = objConn.Execute (ssr) 
  'response.write ssr
   kira = rsssr ("kira")	
  
%>


<table align="center" border="0" cellpadding="0" cellspacing="0" width="98%">
<tr>
  <td align="center" width="10%" valign="top" rowspan="3" >
  <img border="0" src="images/logompsp.jpg"></td>
  <td valign="top" align="center">
      <strong><font size="3" face="Verdana">
      MAJLIS PERBANDARAN SEBERANG PERAI<br>
      SENARAI PEKERJA PERMOHONAN KERJA LEBIH MASA <br>
      BULAN <%=bln%>/<%=thn%><br>
      JABATAN&nbsp;<%=ket%>
      </font></strong></td>
  <td rowspan="3" width="10%"></td>
</tr>


<tr>
  <td colspan=3>&nbsp;</td>
</tr>
</table>


<br><br>
<table align="center" border="0" cellpadding="0" cellspacing="0" width="78%">
<tr>
  <td width="28%" ><font face="Verdana" size="2"><b>&nbsp;Arahan Kerja</b></font></td>
  <td width="2%" ><font face="Verdana" size="2"><b>&nbsp;:</b></font></td>
  <td width="70%" ><font face="Verdana" size="2"><input type="hidden" value="<%=id%>" name="id"><%=id%></font></td>
</tr>


<tr>
  <td width="28%" ><font face="Verdana" size="2"><b>&nbsp;Bilangan Pekerja</b></font></td>
  <td width="2%" ><font face="Verdana" size="2"><b>&nbsp;:</b></font></td>
  <td width="70%" ><font face="Verdana" size="2"><%=kira%></font></td>
</tr>
<tr>
  <td width="28%" ><font face="Verdana" size="2"><b>&nbsp;Keterangan Kerja</b></font></td>
  <td width="2%" ><font face="Verdana" size="2"><b>&nbsp;:</b></font></td>
  <td width="70%" ><font face="Verdana" size="2"><% if keterangan <>"" then %><%=keterangan%><%else%><%=keterangan_1%><%end if%>
   <input type="hidden" name="keterangan" value="<%=keterangan%>"></font></td>
</tr>
<tr>
  <td width="28%" ><font face="Verdana" size="2"><b>&nbsp;Nota Tambahan</b></font></td>
  <td width="2%" ><font face="Verdana" size="2"><b>&nbsp;:</b></font></td>
  <td width="70%" ><font face="Verdana" size="2"><%=nota%><input type="hidden" name="nota" value="<%=nota%>"></font></td>
</tr>

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
    <td width="15%" align="center" bgcolor="#CCCCCC"><b>No Pekerja</b></td>
    <td width="28%" align="center" bgcolor="#CCCCCC"><b>Nama Pekerja</b></td>
    <td width="16%"  align="center" bgcolor="#CCCCCC"><b>Tarikh OT</b></td>
    <td width="18%" align="center" bgcolor="#CCCCCC"><b>Masa Mula (HH24:MM)</b></td>
    <td width="20%" align="center" bgcolor="#CCCCCC"><b>Masa Tamat (HH24:MM)</b></td>
  </tr>


<%
  id = rsp ("id_ak")
  bil = 0
  Do While Not rsss.eof 
  no_pekerja = rsss ("no_pekerja")
  anggaran = 0
   
  n = " select no_pekerja , initcap(nama) nama , jawatan from payroll.paymas " 
  n = n & " where no_pekerja = '"&no_pekerja&"' " 
  Set rsn = objConn.Execute (n)
	nama = rsn ("nama")
	jaw = rsn ("jawatan")
  
  qq = " select to_char(masa_mula_angg,'HH24:MI') as masa_mula , to_char(masa_tamat_angg,'HH24:MI') as masa_tamat ,to_char(tkh_ot,'dd/mm/yyyy') as tkh_ot , "
  qq = qq & " no_pekerja , angg_ot_s , angg_ot_m  "
  qq = qq & " from payroll.pekerja_ot " 
  qq = qq & " where id_ak = '"&id&"' and no_pekerja ='"&no_pekerja&"' "
  qq =qq & " order by tkh_ot, masa_mula , masa_tamat asc  " 
  Set rsqq = objConn.Execute (qq)
   'response.write qq
   Do While Not rsqq.eof 
    bil = bil + 1
	mula = rsqq ("masa_mula")
	tamat = rsqq ("masa_tamat")
	tkh_ot = rsqq ("tkh_ot")
	ot_s = rsqq ("angg_ot_s")
	ot_m = rsqq ("angg_ot_m")
	no_pekerja = rsqq ("no_pekerja")

	anggaran = cdbl(ot_s) +  cdbl(ot_m)

%>

   <tr>
   <td align="center" height="40">&nbsp;<%=no_pekerja%></td>
   <td align="left" height="40">&nbsp;<%=nama%></td>
   <td align="center" height="40">&nbsp;<%=tkh_ot%></td>
   <td align="center" height="40">&nbsp;<%=mula%></td>
   <td align="center" height="40">&nbsp;<%=tamat%></td>
   </tr> 
   <%  
    rsqq.movenext
	loop 
	rsss.movenext
	loop 
   %>
</table>           
<br><br>


<table align="center" border="0" cellpadding="0" cellspacing="0" width="68%">
<br>

<tr>
  <td width="30%" ><font face="Verdana" size="2"><b>Pegawai Yang Meluluskan</b></font></td>
  <td width="1%" ><font face="Verdana" size="2"><b>:</b></font></td>
  <td width="69%" height="70" ><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br><br><br><br>
    <%=nama3%> - (<%=penyelia%>)</font></td>
</tr>
<tr>
  <td width="30%"></td>
  <td width="1%"></td>
  <td width="69%" height="30" ><font face="Verdana" size="2"><%=jaw3%>
</font></td>
</tr>

<tr><td colspan="3">&nbsp;</td></tr>
<tr><td colspan="3">&nbsp;</td></tr>
<tr><td colspan="3">&nbsp;</td></tr>
<tr><td colspan="3">&nbsp;</td></tr>
</table>


<table align="center" border="0" cellpadding="0" cellspacing="0" width="70%">
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
<tr><td colspan="3">&nbsp;</td></tr>
</table>




