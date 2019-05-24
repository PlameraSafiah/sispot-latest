<html>
<head>
<title>Cetak Arahan Kerja</title>
</head>
<body>



<form action="sp26a.asp" method="POST">

<%Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"

    id = Request.querystring("id_ak")
	pemohon = Request.querystring("pemohon")
    tkh_ot = Request.querystring("tkh_ot")
 
 
 ss = " select id_ak ,pemohon, penyelia , unit , ptj , keterangan_1 ,keterangan, nota , tahun , bulan"
 ss = ss & " from payroll.arahan_kerja where id_ak='"&id&"' " 
  Set rsss = objConn.Execute (ss)
  'response.write ss
  	id = rsss ("id_ak")
	penyelia = rsss ("penyelia")
	keterangan = rsss("keterangan")
        keterangan_1 = rsss("keterangan_1")  
	pemohon = rsss ("pemohon")
    bln = rsss ("bulan")
	thn = rsss ("tahun")
	unit = rsss ("unit")
    ptj = rsss ("ptj")
	nota = rsss ("nota")
	
	
  p = "  select id_ak , no_pekerja , kumpulan  "
  p = p & "  from payroll.pekerja_ot where id_ak='"&id&"' " 
  Set rsp = objConn.Execute(p)
  'response.write p
  id_ak = rsp ("id_ak")
  no_pekerja = rsp ("no_pekerja")
  kumpulan = rsp ("kumpulan")
	
	

	
  se= " select kod , keterangan from iabs.jabatan "
  se= se & " where kod = '"&ptj&"' "
  Set rse = objConn.Execute(se)
  ket = rse ("keterangan")
	
  
   
  n = " select no_pekerja , nama , jawatan from payroll.paymas " 
  n = n & " where no_pekerja = '"&pemohon&"' " 
  Set rsn = objConn.Execute (n)
	nama = rsn ("nama")
	jaw = rsn ("jawatan")


ss = " select no_pekerja, lokasi from payroll.unit_sispot_pelulus "
ss = ss & " where lokasi = '"& ptj &"' " 
set rss1 = objConn.execute(ss)

if not rss1.eof then 
pelulus = rss1("no_pekerja") 
end if

	
  pp = " select no_pekerja , nama , jawatan from payroll.paymas " 
  pp = pp & " where no_pekerja = '"& pelulus &"' " 
  Set rspp = objConn.Execute (pp)
if not rspp.eof then
	nama1 = rspp ("nama")
	jaw1 = rspp ("jawatan")
end if

 qq = " select no_pekerja , nama , jawatan from payroll.paymas " 
  qq= qq & " where no_pekerja = '"& penyelia&"' " 
  Set rsqq = objConn.Execute (qq)
if not rsqq.eof then
	nama11 = rsqq ("nama")
	jaw11 = rsqq ("jawatan")
end if

%>


<table align="center" border="0" cellpadding="0" cellspacing="0" width="98%">
<tr>
  <td align="center" width="10%" valign="top" rowspan="3" >
  <img border="0" src="images/logompsp.jpg"></td>
  <td valign="top" align="center">
      <strong><font size="3" face="Verdana">
      MAJLIS PERBANDARAN SEBERANG PERAI<br>
      ARAHAN KERJA LEBIH MASA<br>
      JABATAN&nbsp;<%=ket%><br>
      NO RUJUKAN :<%=id%><br></font></strong></td>
  <td rowspan="3" width="10%"></td>
</tr>


<tr>
  <td colspan=3>&nbsp;</td>
</tr>
</table>
<br><br>


<table align="center" border="0" cellpadding="0" cellspacing="0" width="78%">
<tr>
  <td width="33%"><font face="Verdana" size="2"><b>&nbsp;Arahan Kerja</b></font></td>
  <td width="2%"><font face="Verdana" size="2"><b>&nbsp;:</b></font></td>
   <td width="65%"><input type="hidden" value="<%=id%>" name="id"><%=id%></td>
</tr>

<tr>
  <td width="33%"><font face="Verdana" size="2"><b>&nbsp;Bulan / Tahun</b></font></td>
  <td width="2%"><font face="Verdana" size="2"><b>&nbsp;:</b></font></td>
  <td><input type="hidden" name="bln" size="2" maxlength="2" value="<%=bln%>"><%=bln%>
        &nbsp;/&nbsp;
      <input type="hidden" value="<%=thn%>" name="thn" size="4" maxlength="4"><%=thn%></td>
</tr>

<tr>
  <td ><font face="Verdana" size="2"><b>&nbsp;Jabatan / Unit</b></font></td>
  <td ><font face="Verdana" size="2"><b>&nbsp;:</b></font></td>
  <td ><font face="Verdana" size="2"><%=ptj%>-<%=rse("keterangan")%> (<%=unit%>)</font></td>
</tr>

<tr>
  <td ><font face="Verdana" size="2"><b>&nbsp;Nama Pemohon </b></font></td>
  <td ><font face="Verdana" size="2"><b>&nbsp;:</b></font></td>
  <td ><font face="Verdana" size="2"><%=pemohon%>-(<%=nama%>) <input type="hidden" name="nama" value="<%=nama%>"></font></td>
</tr>


<tr>
  <td ><font face="Verdana" size="2"><b>&nbsp;Jawatan</b></font></td>
  <td ><font face="Verdana" size="2"><b>&nbsp;:</b></font></td>
  <td ><font face="Verdana" size="2"><%=jaw%></font></td>
</tr>


<tr>
  <td ><font face="Verdana" size="2"><b>&nbsp;Nama Penyelia</b></font></td>
  <td ><font face="Verdana" size="2"><b>&nbsp;:</b></font></td>
  <td ><font face="Verdana" size="2"><%=penyelia%>-(<%=nama11%>)</font></td>
</tr>


<tr>
  <td ><font face="Verdana" size="2"><b>&nbsp;Kumpulan</b></font></td>
  <td ><font face="Verdana" size="2"><b>&nbsp;:</b></font></td>
  <td ><font face="Verdana" size="2"><%=kumpulan%></font></td>
</tr>

<tr>
  <td ><font face="Verdana" size="2"><b>&nbsp;Keterangan Kerja</b></font></td>
  <td ><font face="Verdana" size="2"><b>&nbsp;:</b></font></td>
  <td ><font face="Verdana" size="2"><% if keterangan <>"" then %><%=keterangan%><%else%><%=keterangan_1%><%end if%></font></td>
</tr>

<tr>
  <td ><font face="Verdana" size="2"><b>&nbsp;Nota Tambahan</b></font></td>
  <td ><font face="Verdana" size="2"><b>&nbsp;:</b></font></td>
  <td ><font face="Verdana" size="2"><%=nota%></font></td>
</tr>


<tr><td colspan="3">&nbsp;</td></tr>
<tr><td colspan="3">&nbsp;</td></tr>
</table><br><br>

<table align="center" border="0" cellpadding="0" cellspacing="0" width="88%">
<tr>
  <td align="center" valign="top" colspan="3">
  </td>
</tr>
<tr>
  <td colspan="3">Dengan ini tuan/puan diarah menjalankan kerja lebih masa untuk disiapkan/dilaksanakan dalam tempoh yang ditetapkan seperti berikut : </td></tr>
</table><br><br>

<table align="center" border="0" cellpadding="0" cellspacing="0" width="88%">
  <tr class="hd"> 
    <td width="3%" align="center">Bil</b></td>
    <td width="23%" align="center">Tarikh</b></td>
    <td width="18%" align="center">Masa Mula (HH24:MM)</td>
    <td width="22%" align="center">Masa Tamat (HH24:MM)</td>
    <td width="10%" align="center">Hari</td>
    <td width="24%" align="center">Anggaran OT (RM)</td> 
  </tr>
</table>
<hr width="88%">


<table align="center" border="0" cellpadding="0" cellspacing="0" width="88%">
<%   
   id = rsss ("id_ak")	
   pemohon = rsss ("pemohon")
   bil = 0
   
  qq = " select to_char(masa_mula,'HH24:MI') as masa_mula ,to_char(masa_tamat,'HH24:MI') as masa_tamat , to_char(tkh_ot,'dd/mm/yyyy') as tkh_ot , kategori_hari "
  qq = qq & " from payroll.jadual_ot " 
  qq = qq & " where id_ak = '"&id&"' order by tkh_ot , masa_mula ,masa_tamat asc  " 
  Set rsqq = objConn.Execute (qq)
  'response.write qq
  Do While Not rsqq.eof 
    bil = bil + 1
    tkh_ot = rsqq ("tkh_ot")
    mula = rsqq ("masa_mula")
	tamat = rsqq ("masa_tamat")
	hari = rsqq ("kategori_hari")
 %>
 
   <tr>
   <td width="3%" align="center">&nbsp;<%=bil%></td>
   <td width="23%" align="center">&nbsp;<%=tkh_ot%></td>
   <td width="18%" align="center">&nbsp;<%=mula%></td>
   <td width="22%" align="center">&nbsp;<%=tamat%></td>
   <td width="10%" align="center">&nbsp;<%=hari%></td>
   
   
  <%   
   id = rsss ("id_ak")	
   pemohon = rsss ("pemohon")
   tkh_ot = rsqq ("tkh_ot")
   mula = rsqq ("masa_mula")
   tamat = rsqq ("masa_tamat")
   
   se = " select sum(nvl(angg_ot_s,0) + nvl(angg_ot_m,0)) anggaran "
   se = se & " from payroll.pekerja_ot "
   se = se & " where id_ak='"&id&"' and to_char(tkh_ot,'dd/mm/yyyy') ='"&tkh_ot&"' and  to_char(masa_mula_angg,'HH24:MI')= '"&mula&"' "
   se = se & " and  to_char(masa_tamat_angg,'HH24:MI')= '"&tamat&"'   " 
   Set rse= objConn.Execute (se)
   anggaran = rse("anggaran")  
%>    

   <td width="24%" align="center">&nbsp;<%=anggaran%>
    </td>    
   </tr>
   
   
<%     
 rsqq.movenext
 loop 
%>


</table>
<hr width="88%">


  <%
   sss = " select sum(nvl(angg_ot_s,0) + nvl(angg_ot_m,0)) semuaxgst1 "
   sss = sss & " from payroll.pekerja_ot"
   sss = sss & " where id_ak = '"&id&"'  " 
   Set rssss = objConn.Execute (sss)
    'response.write sss
   semuaxgst1 = rssss("semuaxgst1")
  %> 
  
<table align="center" border="0" cellpadding="0" cellspacing="0" width="88%">
<tr>
 <td colspan="2" align="right"><font face="Verdana" size="2"><B>Jumlah Anggaran OT (RM)</B></font></td>
 <td width="24%" align="center"><font face="Verdana" size="2"><B><%=semuaxgst1%></B></font></td>
</tr>
</table>
<hr width="88%">

<br><br><br>

<table align="center" border="0" cellpadding="0" cellspacing="0" width="93%">
<tr>
  <td width="49%" ><font face="Verdana" size="2"><b>&nbsp;Pegawai Yang Memohon</b></font></td>
  <td width="51%" ><font face="Verdana" size="2"><b>&nbsp;Pegawai Yang Meluluskan</b></font></td>
</tr>
</tr>

<tr>
  <td height="50"><font face="Verdana" size="2"><b>&nbsp;Tandatangan/Cop : </b></font></td>
  <td height="50"><font face="Verdana" size="2"><b>&nbsp;Tandatangan/Cop : </b></font></td>
</tr>
<tr>
  <td height="30"><font face="Verdana" size="2"><b>&nbsp;Nama : </b><%=nama%></font></td>
  <td height="30"><font face="Verdana" size="2"><b>&nbsp;Nama :</b> <%=nama11%></font></td>
</tr>
<tr>
  <td height="40"><font face="Verdana" size="2"><b>&nbsp;Jawatan :</b> <%=jaw%></font></td>
  <td height="40"><font face="Verdana" size="2"><b>&nbsp;Jawatan :</b> <%=jaw11%></font></td>
</tr>
<tr>
  <td><font face="Verdana" size="2"><b>&nbsp;Tarikh : </b></font></td>
  <td><font face="Verdana" size="2"><b>&nbsp;Tarikh : </b></font></td>
</tr>
</table>