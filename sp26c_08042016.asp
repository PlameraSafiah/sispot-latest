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
  p = "  select id_ak , penyelia , pemohon, pengesahan, tkh_sah, unit, ptj, keterangan, nota, tkh_input, tahun, bulan "
  p = p & "  from payroll.arahan_kerja where id_ak='"&id&"' "
  Set rsp = objConn.Execute(p)
  'response.write p
  id = rsp ("id_ak")
  bln = rsp ("bulan")
  thn = rsp ("tahun")
  pemohon = rsp ("pemohon")
  penyelia = rsp ("penyelia")
  keterangan = rsp ("keterangan")
  nota = rsp ("nota")
  ptj = rsp ("ptj")
  
  
  'bil pekerja
  ssr = " select count(distinct no_pekerja) as kira"
  ssr = ssr & " from payroll.pekerja_ot " 
  ssr = ssr & " where id_ak = '"&id&"' order by no_pekerja" 
  Set rsssr = objConn.Execute (ssr) 
  'response.write ssr
  kira = rsssr ("kira")
  
  
  'nama jabatan
  m = " select kod , keterangan from iabs.jabatan "
  m = m & " where kod = '"&ptj&"' "
  Set rsm = objConn.Execute(m)
  ket = rsm ("keterangan")
  
  
   'nama penyelia
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
  <td width="70%" ><font face="Verdana" size="2"><%=keterangan%><input type="hidden" name="keterangan" value="<%=keterangan%>"></font></td>
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


<%
  id = rsp ("id_ak")
  bil = 0
  anggaran = 0
  
  qq = " select distinct to_char(tkh_ot,'dd/mm/yyyy') as tkh_ot"
  qq = qq & " from payroll.pekerja_ot " 
  qq = qq & " where id_ak = '"&id&"'  "
  qq =qq & " order by tkh_ot  " 
  Set rsqq = objConn.Execute (qq)
   'response.write qq
   Do While Not rsqq.eof 
   bil = bil + 1
   tkh_ot = rsqq ("tkh_ot")
	
	
  qqq = " select distinct to_char(masa_mula_angg,'HH24:MI') as masa_mula , to_char(masa_tamat_angg,'HH24:MI') as masa_tamat "
  qqq = qqq & " from payroll.pekerja_ot " 
  qqq = qqq & " where id_ak = '"&id&"' and to_char(tkh_ot,'dd/mm/yyyy') = '"&tkh_ot&"'  "
  Set rsqqq = objConn.Execute (qqq)
   'response.write qqq
  mula = rsqqq ("masa_mula")
  tamat = rsqqq ("masa_tamat")
%>



<table align="center" border="1" cellpadding="0" cellspacing="0" width="70%">
  <tr class="hd"> 
    <td width="16%" align="left" bgcolor="#CCCCCC" colspan="2"><b>Tarikh OT : &nbsp;<%=tkh_ot%></b>
     &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
      &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
      &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
     <b>Masa OT (HH24:MM) : &nbsp;<%=mula%> - <%=tamat%></b>
    </td>
  </tr>      
  <br>
         
         
<%
  'no_pekerja - jadual
  ss = " select distinct no_pekerja "
  ss = ss & " from payroll.pekerja_ot " 
  ss = ss & " where id_ak = '"&id&"' and to_char(tkh_ot,'dd/mm/yyyy') = '"&tkh_ot&"'  and to_char(masa_mula_angg,'HH24:MI')= '"&mula&"' and "
  ss = ss & " to_char(masa_tamat_angg,'HH24:MI') = '"&tamat&"' order by no_pekerja " 
  Set rsss = objConn.Execute (ss) 
  Do While Not rsss.eof 
   'response.write ss
   no_pekerja = rsss ("no_pekerja")	
 
  n = " select no_pekerja , initcap(nama) nama , jawatan from payroll.paymas " 
  n = n & " where no_pekerja = '"&no_pekerja&"' " 
  Set rsn = objConn.Execute (n)
  nama = rsn ("nama")
  jaw = rsn ("jawatan")
  
%>

    <tr> 
    <td width="16%" align="left" colspan="2">
    &nbsp;<%=no_pekerja%> - <%=nama%><br>
    </td>
   </tr>  
    <%
	rsss.movenext
	loop
    rsqq.movenext
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




