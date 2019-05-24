<%'<--#INCLUDE FILE="adovbs.inc" -->%>
<!--#include file="connection.asp" -->
<%response.buffer=true%>
<%response.cookies("amenu") = "sp27b.asp"%>
<!--'#include file="spmenu.asp"-->


<html>
<head>
<title>Semakan Status Tuntutan OT Mengikut Pekerja</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" href="css/menu1.css" REL="stylesheet">
<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%=color4%>">
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
<script language="javascript">
function submitForm1()
{
	document.test.submit();
}
</script>
</head>


<body leftmargin="0" topmargin="0" OnLoad="cssHorScrolls();">
<%response.cookies("amenu") = "sp27b.asp"%>


<% 
proses = request.form("proses")
tkh_ot = request.form("tkh_ot")
ptj = request.form("ptj")
pemohon = request.form("pemohon")
blnthn = request.form("blnthn")
id_ak = request.form("id_ak")
p7 = request.form("b7")	
ptj=plokasi


papar
if proses <> "" then	
	borang
end if


sub papar ()
if blnthn = "" then
Dim TodayYearNumber, TodayMonthName
TodayYearNumber = DatePart("yyyy",Date)
TodayMonth = DatePart("m",Date)
if len(TodayMonth)="1" then TodayMonth="0"& TodayMonth
blnthn=TodayMonth&TodayYearNumber
end if


	'check jabatan mengikut id login ---> nadia (05092016)
	pekd = request.cookies("gnop")
	a =     "select jabatan from majlis.pengguna where no_pekerja = '"& pekd &"' "
	Set rsa = objConn.Execute(a)

    'semakan id login
 	idd = "select a.nama, a.lokasi, b.keterangan, to_char(sysdate,'dd/mm/yyyy') tkh  from payroll.paymas a, iabs.jabatan b "
 	idd = idd & " where a.no_pekerja = '"&pekd&"' and b.kod = a.lokasi "
 	set idd2 = objConn.Execute(idd)
 	lokasi = idd2 ("lokasi")
 
%>
<br>
<br>


<form name="test" method="post" action="sp27b.asp" onSubmit="return check(this)">
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
    <tr class="hd">
    <td colspan="2">Semakan Status Tuntutan OT Mengikut Pekerja</td>
    </tr>    
    <tr>
  	<td width="15%" align="right"><b>Tahun &nbsp;&nbsp;</b></td>
  	<td><!--<input type="text" name="tkh_ot" size="8" value="<%=tkh_ot%>">-->
    <select size="1" name="tkh_ot" onChange="submitForm1(this)">
    <option value="">Pilih Tahun</option>
    
	<% 
    'filter by tarikh ot ---> nadia (16032016)
	q2="select distinct to_char(tkh_ot,'yyyy') as tkh_ot , ptj "
	q2 = q2 & " from payroll.proses_ot where ptj = '"&lokasi&"' "
	q2 = q2 & " order by tkh_ot desc"  
	set oq2 = objConn.Execute(q2)
	if not oq2.bof and oq2.eof then
	tkh_ot = oq2 ("tkh_ot")
	ptj = oq2 ("ptj")
	end if
			  
    do while not oq2.EOF %>
    <option <%if tkh_ot = oq2("tkh_ot") then%> selected <%end if%>value="<%=oq2("tkh_ot")%>"><%=oq2("tkh_ot")%></option>
     <%
     oq2.MoveNext
     loop
'------
oq2.close
'------
     %>
    </select>
    </td>
    </tr> 
     
   <tr>
   <td align="right"><b>Nama Pekerja &nbsp;&nbsp;</b></td>
   <td><!--<input type="text" name="pemohon" size="8" value="<%=pemohon%>">-->
   <select size="1" name="pemohon">
   <option value="">Pilih Pekerja</option>
   
   <% 	  
   'filter by pekerja ---> nadia (16032016)
   if tkh_ot <> "" then 
   'listkan semua no pekerja mengikut bulan ot , yang telah selesai proses di kew ---> nadia (05092016)
   q2b=" select distinct a.no_pekerja, to_char(a.tkh_ot,'yyy') , b.nama "
   q2b = q2b & " from payroll.proses_ot a, payroll.paymas b where a.ptj = '"&ptj&"' and to_char(a.tkh_ot,'yyyy') = '"&tkh_ot&"' "
   q2b = q2b & " and a.no_pekerja = b.no_pekerja order by a.no_pekerja asc"
   set oq2b = objConn.Execute(q2b)
   if not oq2b.bof and oq2b.eof then
   pemohon = oq2b ("no_pekerja")
   nama = oq2b ("nama")
   end if
   
   do while not oq2b.EOF %>
   <option <%if cstr(pemohon) =  cstr(oq2b("no_pekerja")) then%> selected <%end if%>value="<%=oq2b("no_pekerja")%>"><%=oq2b("no_pekerja")%> - <%=oq2b("nama")%></option>
   <%
   oq2b.MoveNext
   loop	
'------
oq2b.close
'------	
   end if 
   %>
   </select>
   </td>
   </tr>  
   <tr align="center"><td colspan="2"><input type="submit" name="proses" value="Hantar"></td>
   </tr>
</table>
</form>
<% end sub %>



<%
sub borang()
blnthn=request.form("blnthn")
tkh_ot = request.form("tkh_ot")
pemohon = request.form("pemohon")
ptj = plokasi

	
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


   
	'keterangan jabatan mengikut jabatan pekerja
     m = " select kod , initcap (keterangan) keterangan from iabs.jabatan "
     m = m & " where kod = '"&ptj&"' "
     Set rsm = objConn.Execute(m)
     ket = rsm ("keterangan")	
	 
	 
	'nama pekerja
     md = " select nama from payroll.paymas "
     md = md & " where no_pekerja = '"&pemohon&"' "
     Set rsmd = objConn.Execute(md)
     nam = rsmd ("nama")	
%>
<br>
<br>


<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd" height="20%">
  <TBODY>
    <tr class="hd"> 
    <td colspan="10">Semakan Status Tuntutan Lebih Masa : &nbsp;&nbsp;<b><%=nam%></b><br>Jabatan <%=ket%><br>Tuntutan : <%=tkh_ot%></td></tr> 
    <tr style="BACKGROUND-COLOR: #ffffff"><td colspan="10">&nbsp;</td></tr>
    <tr class="hd"> 
    <td width="2%"   align="center">Bil</td>
    <td width="4%"   align="center">Bulan</td>
    <td width="6%"   align="center">Tahun</td>
    <td width="9%"   align="center">No Kawalan</td>
    <td width="27%"  align="center">Nama Pekerja</td>
    <td width="7%"   align="center">Jabatan</td>
    <td width="9%"   align="center">Tarikh Sah Kewangan</td>
    <td width="10%"  align="center">Status Pengesahan</td>
    <td width="14%"  align="center">Anggaran OT (RM)</td>
    <td width="12%"  align="center">Bayaran OT Sebenar (RM)</td>
    </tr>
  
  
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
bil=bil + 1


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
	'response.write se
	if not rse.bof and rse.eof then
   	anggaran = rse("amaun")
	end if


%>



<form name="test<%=bil%>"  method="post" action="sp27b.asp" target="_new">
  <tr bgcolor="#EFF5F5">
    <td rowspan="2" align="center"><%=bil%></td>
    <td rowspan="2" align="center"><%=rssss("bln")%></td>
    <td rowspan="2" align="center"><%=rssss("thn")%></td>
    <td rowspan="2" align="center"><%=rssss("no_kawalan")%></td>
    <td rowspan="2" align="left"><%=rssss("no_pekerja")%> - <%=rsmd ("nama")%></td>
    <td rowspan="2" align="center"><%=ptj%></td>
    <td rowspan="2" align="center"><%=rssss("tkh_sah_kew")%></td>
    <td rowspan="2" align="center"> 
	<%if sah_kew="Y" then%>Y<%end if%>
	<%if sah_kew="T" then%><font color="#e60000"><b>Belum Disahkan</b></font><%end if%>
    </td>
   <td rowspan="2" align="center"><%=formatnumber(rsss("tuntutan"),2)%></td>
   <td rowspan="2" align="center">
   <%if sah_kew="Y" then%><%=formatnumber(rse("amaun"),2)%><%else%> <%end if%>
  </td>
  </tr>
 </form>
 <form name="test<%=bil%>"  method="post" action="sp27b.asp" target="_new">
  <tr bgcolor="#EFF5F5">
  </tr>
 </form>  

<%
rssss.movenext
loop
'-----
rssss.close
'-----
%>

</tbody> 
</table>
<br>

<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
  <tr align="center">
  <td colspan="2">
  <form name="test<%=bil%>"  method="post" action="sp27b1.asp?tkh_ot=<%=thn%>&pemohon=<%=no_pekerja%>&amenu='sp27b1'" target="_new">
  <input type="submit" name="b7" value="Cetak Laporan" class="button" onFocus="nextfield='done';">
  </form>
  </td>
  </tr>
</table>
	


<% 
'----
objConn.close
 '----- 
end sub %>
<input type="hidden" name="bilrec" value="<%=ctrz%>" >
</body>
</html>