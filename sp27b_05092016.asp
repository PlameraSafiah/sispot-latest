<%'<--#include file="connection.asp" -->
'<--#INCLUDE FILE="adovbs.inc" -->%>
<!--#include file="connection.asp" -->
<%response.buffer=true%>
<%response.cookies("amenu") = "sp27b.asp"%>
<!--'#include file="spmenu.asp"-->
<html>
<head>
<title>Laporan Semakan Pengesahan Mengikut Bulan & Pemohon</title>
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
<%'<--#include file="banner.asp"-->
'<--#include file="subm.asp"-->
%>


<% 
proses = request.form("proses")
tkh_ot = request.form("tkh_ot")
ptj = request.form("ptj")
pemohon = request.form("pemohon")
blnthn = request.form("blnthn")
id_ak = request.form("id_ak")
p7 = request.form("b7")	'cetak permohonan
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



	'check jabatan by id login
	pekd = request.cookies("gnop")
	a =     "select jabatan from majlis.pengguna where no_pekerja = '"& pekd &"' "
	Set rsa = objConn.Execute(a)

    'details id login
 	idd = "select a.nama, a.lokasi, b.keterangan, to_char(sysdate,'dd/mm/yyyy') tkh  from payroll.paymas a, iabs.jabatan b "
 	idd = idd & " where a.no_pekerja = '"&pekd&"' and b.kod = a.lokasi "
 	set idd2 = objConn.Execute(idd)
 	'response.write idd
 	lokasi = idd2 ("lokasi")
 
%>
<br>
<br>


<form name="test" method="post" action="sp27b.asp" onSubmit="return check(this)">
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
    <tr class="hd">
      <td colspan="2">Laporan Arahan Kerja - Semakan Pengesahan Arahan Kerja Mengikut Bulan &amp; Pemohon</td>
    </tr>
    
    <tr>
  	<td width="15%"><b>Bulan/Tahun (mm/yyyy) : </b></td>
  	<td><!--<input type="text" name="tkh_ot" size="8" value="<%=tkh_ot%>">-->
     <select size="1" name="tkh_ot" onChange="submitForm1(this)">
          <option value="">Pilih Bulan/Tahun</option>
    <% 
     'filter by tarikh ot ---> nadia (16032016)
	q2="select distinct to_char(a.tkh_ot,'mm/yyyy') as tkh_ot , b.ptj "
	q2 = q2 & " from payroll.jadual_ot a , payroll.arahan_kerja b where a.id_ak=b.id_ak and b.ptj = '"&lokasi&"' "
	q2 = q2 & " order by tkh_ot asc"  
	set oq2 = objConn.Execute(q2)
	'response.write q2
	if not oq2.bof and oq2.eof then
	tkh_ot = oq2 ("tkh_ot")
	ptj = oq2 ("ptj")
	end if
		  
		  do while not oq2.EOF %>
          <option <%if tkh_ot = oq2("tkh_ot") then%> selected <%end if%>value="<%=oq2("tkh_ot")%>"><%=oq2("tkh_ot")%></option>
          <%
            oq2.MoveNext
            loop
          %>
    </select>
    </td>
    </tr> 
    
    
   <tr>
   <td><b>Nama Pemohon : </b></td>
   <td><!--<input type="text" name="pemohon" size="8" value="<%=pemohon%>">-->
   <select size="1" name="pemohon">
   <option value="">Pilih Pemohon</option>
   <% 	  
  'filter by pemohon ---> nadia (16032016)
   if tkh_ot <> "" then 
   q2b=" select distinct a.pemohon, to_char(c.tkh_ot,'mm/yyyy') , b.nama "
   q2b = q2b & " from payroll.arahan_kerja a, payroll.jadual_ot c , payroll.paymas b where a.ptj = '"&ptj&"' and to_char(c.tkh_ot,'mm/yyyy') like '"&tkh_ot&"' "
   q2b = q2b & " and a.id_ak = c.id_ak and a.pemohon = b.no_pekerja order by a.pemohon asc"
   set oq2b = objConn.Execute(q2b)
   'response.write  q2b
   pemohon = oq2b ("pemohon")
   nama = oq2b ("nama")

		  do while not oq2b.EOF %>
          <option <%if cstr(pemohon) =  cstr(oq2b("pemohon")) then%> selected <%end if%>value="<%=oq2b("pemohon")%>"><%=oq2b("pemohon")%> - <%=oq2b("nama")%></option>
          <%
            oq2b.MoveNext
            loop		
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
id_ak = request.form("id_ak")
pemohon = request.form("pemohon")
ptj = plokasi


    'papar data ---> nadia (16032016)
	sss = " select distinct b.id_ak , to_char(b.tkh_ot,'mm/yyyy') as tkh_ot , a.ptj "
	sss = sss & " from payroll.jadual_ot b , payroll.arahan_kerja a  where a.id_ak=b.id_ak and  to_char(b.tkh_ot,'mm/yyyy') ='"& tkh_ot&"' and a.ptj= '"&ptj&"' "
	sss = sss & " and a.pemohon='"&pemohon&"' order by id_ak asc"
	'response.write sss
	Set rssss = objConn.Execute (sss)
	tkh_ot = rssss ("tkh_ot")
	id_ak = rssss ("id_ak")

     'keterangan jabatan
     m = " select kod , initcap (keterangan) keterangan from iabs.jabatan "
     m = m & " where kod = '"&ptj&"' "
     Set rsm = objConn.Execute(m)
     ket = rsm ("keterangan")	
%>
<br>
<br>


<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd" height="20%">
  <TBODY>
    <tr class="hd"> 
      <td colspan="10">Senarai Semakan Status Pengesahan Arahan Kerja <br> Jabatan <%=ket%> <br> Bulan <%=tkh_ot%></td>
    </tr> 
    <tr style="BACKGROUND-COLOR: #ffffff"><td colspan="9">&nbsp;</td></tr>
    <tr class="hd"> 
    <td width="2%"   align="center">Bil</td>
    <td width="9%"   align="center">No Rujukan Arahan Kerja</td>
    <td width="22%"  align="center">Nama Pemohon</td>
    <td width="9%"   align="center">Tarikh Sah</td>
    <td width="6%"   align="center">Jabatan</td>
    <td width="9%"   align="center">Pengesahan</td>
    <td width="9%"   align="center">Jumlah Pekerja</td>
    <td width="17%"  align="center">Anggaran OT (RM)</td>
    <td width="17%"  align="center">Tuntutan OT Sebenar (RM)</td>
    </tr>
  
  
<%  
bil=0
Do While Not rssss.eof 
tkh_ot = rssss ("tkh_ot")
id_ak = rssss ("id_ak")
bil=bil + 1

   'papar tarikh pengesahan ---> nadia (16032016)
   ss = " select  id_ak , ptj , to_char(tkh_sah,'dd/mm/yyyy') as tkh_sah , pengesahan  , bulan , tahun , pemohon "
   ss = ss & " from payroll.arahan_kerja where "
   ss =ss & "  ptj='"&ptj&"' and id_ak='"&id_ak&"' "
   ss = ss & " order by id_ak , ptj  " 
   'response.write ss
   Set rsss = objConn.Execute (ss)
   id_ak = rsss ("id_ak")
   ptj = rsss ("ptj")
   pengesahan = rsss ("pengesahan")
   tkh_sah = rsss ("tkh_sah")
   pemohon = rsss ("pemohon")

   'nama pemohon
   pp = " select no_pekerja , nama , jawatan from payroll.paymas " 
   pp = pp & " where no_pekerja = '"&pemohon&"' " 
   Set rspp = objConn.Execute (pp)
   nama1 = rspp ("nama")
   jaw1 = rspp ("jawatan")	
   
	'keterangan jabatan
    m = " select kod , initcap(keterangan) keterangan from iabs.jabatan "
    m = m & " where kod = '"&ptj&"' "
    Set rsm = objConn.Execute(m)
    ket = rsm ("keterangan")	


    se = " select sum(nvl(angg_ot_s,0) + nvl(angg_ot_m,0)) anggaran , count(distinct no_pekerja) as kira , sum(nvl(ot_sebenar_m,0) +  nvl(ot_sebenar_s,0)) sebenar "
    se = se & " from payroll.pekerja_ot"
    se = se & " where id_ak='"&id_ak&"'  " 
    Set rse= objConn.Execute (se)
   'response.write se
   Do While Not rse.eof 
   anggaran = rse("anggaran")
   kira = rse("kira")
   sebenar = rse("sebenar")
%>



<form name="test<%=bil%>"  method="post" action="sp27b.asp" target="_new">
  <tr bgcolor="#EFF5F5">
    <td rowspan="2" align="center"><%=bil%></td>
    <td rowspan="2" align="center"><%=rssss ("id_ak")%></td>
    <td rowspan="2" align="left"><%=rsss ("pemohon")%>- <%=rspp ("nama")%></td>
    <td rowspan="2" align="center"><%=rsss ("tkh_sah")%></td>
    <td rowspan="2" align="center"><%=rsss ("ptj")%></td>
    <td rowspan="2" align="center"> 
	<%if pengesahan="Y" then%>Y<%end if%>
	<%if pengesahan="T" then%><font color="#e60000"><b>Belum Disahkan</b></font><%end if%>
    </td>
    <td rowspan="2" align="center"><%=rse ("kira")%></td>
    <td rowspan="2" align="center"><%=rse("anggaran")%></td>
    <td rowspan="2" align="center"><%=rse("sebenar")%></td>
  </tr>
 </form>
 <form name="test<%=bil%>"  method="post" action="sp27b.asp" target="_new">
  <tr bgcolor="#EFF5F5">
  </tr>
 </form>  
 
 
<%  
rse.movenext
loop 
rssss.movenext
loop
%>

 
<%  
   sss = " select sum(nvl(a.angg_ot_s,0) + nvl(a.angg_ot_m,0)) semuaxgst1 , sum(nvl(ot_sebenar_m,0) +  nvl(ot_sebenar_s,0)) sebenar1 "
   sss = sss & " from payroll.pekerja_ot a , payroll.arahan_kerja c "
   sss = sss & " where a.id_ak = c.id_ak and to_char(a.tkh_ot,'mm/yyyy') ='"& tkh_ot &"' and c.ptj='"&ptj&"' and c.pemohon='"&pemohon&"' " 
   Set rssss = objConn.Execute (sss)
   'response.write sss
   semuaxgst1 = rssss("semuaxgst1")
   sebenar1 = rssss("sebenar1")
%> 

   <tr bgcolor="#EFF5F5">
    <td colspan="7" align="right"><B>&nbsp; Jumlah Keseluruhan (RM)&nbsp;</B></td>
    <td width="17%" align="center"><b><%=semuaxgst1%></b></td>
       <td width="17%" align="center"><b><%=sebenar1%></b></td>
    </tr>
  </tbody> 
</table>
<br>

  <TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
    <tr align="center">
    <td colspan="2">
   <form name="test<%=bil%>"  method="post" action="sp27b1.asp?tkh_ot=<%=tkh_ot%>&pemohon=<%=pemohon%>&amenu='sp27b1'" target="_new">
   <input type="submit" name="b7" value="Cetak Laporan" class="button" onFocus="nextfield='done';">
    </form>
    </td>
     </tr>
  </table>
	  

<% end sub %>
<input type="hidden" name="bilrec" value="<%=ctrz%>" >
</body>
</html>