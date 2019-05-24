<html>
<head>
<title>Semakan Status Tuntutan OT Mengikut Pekerja</title>
</head>
<body>
<script type="text/javascript" src="menu1.js"></script>
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
function check(f){
	if (f.blnthn.value=="")
	{
		alert("Sila masukkan tahun arahan kerja");
		f.blnthn.focus();
		return false
	}
}
//  End -->
</script>



<form action="sp27b1.asp" method="POST">
<%Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"


blnthn=request.querystring("blnthn")
tkh_ot=request.querystring("tkh_ot")
pemohon = request.querystring("pemohon")
proses = request.querystring("proses")
ptj = plokasi
 
 
 
 	'check jabatan mengikut id login ---> nadia (05092016)
	pekd = request.cookies("gnop")
	a =     "select jabatan from majlis.pengguna where no_pekerja = '"& pekd &"' "
	Set rsa = objConn.Execute(a)

    'semakan id login
 	idd = "select a.nama, a.lokasi, b.keterangan, to_char(sysdate,'dd/mm/yyyy') tkh  from payroll.paymas a, iabs.jabatan b "
 	idd = idd & " where a.no_pekerja = '"&pekd&"' and b.kod = a.lokasi "
 	set idd2 = objConn.Execute(idd)
 	lokasi = idd2 ("lokasi")
	
   
	
	'papar data pekerja mengikut bulan ot ---> nadia (16032016)
	sss = "select distinct to_char(tkh_ot,'mm') as bln , to_char(tkh_ot,'yyyy') as thn ,  to_char(tkh_ot,'mm/yyyy') as tkh_ot , ptj , no_kawalan,  "
	sss = sss & " no_pekerja , to_char(tkh_sah_kew,'dd/mm/yyyy') as tkh_sah_kew , sah_kew  "
	sss = sss & "from payroll.proses_ot where to_char(tkh_ot,'yyyy') ='"& tkh_ot&"'  "
	sss = sss & " and no_pekerja='"&pemohon&"' and no_kawalan is not null order by bln desc"
	Set rssss = objConn.Execute (sss) 
    if not rssss.bof and rssss.eof then 
	tkh_ot = rssss ("tkh_ot")
	bln = rssss ("bln")
	thn = rssss ("thn")
	no_kawalan = rssss ("no_kawalan")
	no_pekerja = rssss ("no_pekerja")
	tkh_sah_kew = rssss ("tkh_sah_kew")
    sah_kew = rssss ("sah_kew")
	end if


    'keterangan jabatan    
    m = " select kod , keterangan from iabs.jabatan "
    m = m & " where kod = '"&lokasi&"' "
    Set rsm = objConn.Execute(m)
    ket = rsm ("keterangan")	
	
		 
	'nama pekerja
     md = " select nama from payroll.paymas "
     md = md & " where no_pekerja = '"&pemohon&"' "
     Set rsmd = objConn.Execute(md)
     nam = rsmd ("nama")	
  
%>


<table align="center" border="0" cellpadding="0" cellspacing="0" width="98%">
<tr>
  <td align="center" width="10%" valign="top" rowspan="3" >
  <img border="0" src="images/logompsp.jpg"></td>
  <td valign="top" align="center">
      <strong><font size="3" face="Verdana">
      MAJLIS PERBANDARAN SEBERANG PERAI<br>
      STATUS TUNTUTAN LEBIH MASA<br>
      TAHUN <%=tkh_ot%>
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
    <td width="2%"     align="center"><b>Bil</b></td>
    <td width="6%"     align="center"><b>Bulan</b></td>
    <td width="7%"     align="center"><b>Tahun</b></td>
    <td width="8%"     align="center"><b>No Kawalan</b></td>
    <td width="28%"    align="center"><b>Nama Pekerja</b></td>
    <td width="8%"     align="center"><b>Jabatan</b></td>
    <td width="7%"     align="center"><b>Status Pengesahan</b></td>
    <td width="11%"    align="center"><b>Tarikh Sah<br>Kewangan</b></td>
    <td width="11%"    align="center"><b>Anggaran OT <br>(RM)</b></td>
    <td width="12%"    align="center"><b>Bayaran OT<br>(RM)</b></td>
</table>
<hr width="98%">
  
  
  
<%  
bil=0
Do While Not rssss.eof
tkh_ot = rssss ("tkh_ot")
bln = rssss ("bln")
thn = rssss ("thn")
no_kawalan = rssss ("no_kawalan")
no_pekerja = rssss ("no_pekerja")
tkh_sah_kew = rssss ("tkh_sah_kew")
sah_kew = rssss ("sah_kew")
ptj = plokasi
lokasi = idd2 ("lokasi")
bil=bil+ 1


   'papar tarikh pengesahan ---> nadia (16032016)
   ss = " select sum (tuntutan) as tuntutan"
   ss = ss & " from payroll.proses_ot where "
   ss =ss & "  no_pekerja ='"&no_pekerja&"' and no_kawalan='"&no_kawalan&"' and bln='"&bln&"' and thn='"&thn&"' "
   Set rsss = objConn.Execute (ss)
   if not rsss.bof and rsss.eof then
   tuntutan = rsss ("tuntutan")
   end if
   
   
   	se = " select distinct nvl(a.no_kawalan,0) as no_kawalan, nvl(a.bulan,0) as bulan, nvl(a.tahun,0) as tahun, nvl(sum(nvl(a.amaun,0)),0) as amaun,"
   	se = se & "nvl(sum(nvl(a.A1125,0)),0) as satu, nvl(sum(nvl(a.A125,0)),0) as dua, nvl(sum(nvl(a.A150,0)),0) as tiga,"
   	se = se & "nvl(sum(nvl(a.A175,0)),0) as empat, nvl(sum(nvl(a.A2,0)),0) as lima  " 
	se = se & " from payroll.tuntutan_harian a "
	se = se & " where a.no_pekerja="& no_pekerja &" and a.bulan='"&bln&"' and a.tahun='"&thn&"' and a.no_kawalan='"&no_kawalan&"' "
	se = se & " group by a.no_kawalan,a.bulan,a.tahun order by tahun desc,bulan desc"
   	Set rse= objConn.Execute (se)
	if not rse.bof and rse.eof then
   	anggaran = rse("amaun")
	end if
		
%>


<table align="center" border="0" cellpadding="0" cellspacing="0" width="98%">
  <tr class="hd"> 
    <td width="2%"    rowspan="2"  align="center"><%=bil%></td>
    <td width="6%"    rowspan="2"  align="center"><%=rssss("bln")%></td>
    <td width="7%"    rowspan="2"  align="center"><%=rssss("thn")%></td>
    <td width="8%"    rowspan="2"  align="center"><%=rssss("no_kawalan")%></td>
    <td width="28%"   rowspan="2"  align="center"><%=rssss("no_pekerja")%> - <%=rsmd ("nama")%></td>
    <td width="8%"    rowspan="2"  align="center"><%=idd2 ("lokasi")%></td>
    <td width="7%"   rowspan="2"  align="center"> 
	<%if sah_kew="Y" then%>Y<%end if%>
	<%if sah_kew="T" then%><b>T</b><%end if%>
    </td>
    <td width="11%"    rowspan="2"  align="center"><%=rssss("tkh_sah_kew")%></td>
    <td width="11%" rowspan="2"  align="center"><%=formatnumber(rsss("tuntutan"),2)%></td>
    <td width="12%" rowspan="2"  align="center"><%if sah_kew="Y" then%><%=formatnumber(rse("amaun"),2)%><%else%> <%end if%></td>
  </tr>
<%
rssss.movenext
loop
%>
</table>
<hr width="98%">