<!--#include file="connection.asp" -->
<%response.buffer=true%>
<%response.cookies("amenu") = "sp27c.asp"%>
<!--'#include file="spmenu.asp"-->

<html>
<head>
<title>Cetak Senarai Pekerja</title>
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


</head>
<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%'=color4%>">

<% 
proses = request.form("proses")
tkh_ot=request.form("tkh_ot")
pekerja=request.form("pekerja")
pemohon = request.form ("no_pekerja")
blnthn = request.form("blnthn")
p7 = request.form("b7")	
id = request.form("id_ak")


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
  lokasi = idd2 ("lokasi")
 
 
  q2="select distinct to_char(a.tkh_ot,'mm/yyyy') as tkh_ot , b.ptj , b.pengesahan "
  q2 = q2 & " from payroll.jadual_ot a , payroll.arahan_kerja b where a.id_ak=b.id_ak and b.ptj = '"&lokasi&"' and b.pengesahan='Y' order by tkh_ot asc"  
  set oq2 = objConn.Execute(q2)
  if not oq2.bof and oq2.eof then
  tkh_ot = oq2 ("tkh_ot")
  ptj = oq2 ("ptj")
  end if
%>
<br>


<form name="test" method="post" action="sp27c.asp" onSubmit="return check(this)">
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
  <tr class="hd">
    <td colspan="2">Laporan Arahan Kerja - Melebihi Sepertiga Mengikut Bulan</td>
  </tr>
  <tr>
  <td width="15%"><b>Bulan/Tahun(mmyyyy):</b></td>
  <td>       
    <select size="1" name="tkh_ot" value="<%%>" onFocus="nextfield='proses';" >
          <%	if tkh_ot = "" then	%>
          <option value="">Pilih Bulan/Tahun</option>
          <%	else	%>
          <option value="<%=tkh_ot%>"></option>
          <%  end if	
		  
          do while not oq2.EOF 
          %>
          <option value="<%=oq2("tkh_ot")%>"><%=oq2("tkh_ot")%></option>
          <%
            oq2.MoveNext
            loop
'-----
oq2.close
'-----
           %>
     </select>
  </td>
  </tr>
<tr align="center"><td colspan="2"><input type="submit" name="proses" value="Hantar"></td></tr>
</table>
</form>
<%
end sub %>



<%
sub borang()
blnthn=request.form("blnthn")
ptj=plokasi
p7 = request.form("b7")	
tkh_ot=request.form("tkh_ot")


  'list semua pekerja yang membuat ot by bulan --> nadia (01082016)
  ss = " select distinct (a.no_pekerja) ,  to_char(a.tkh_ot,'mm/yyyy') as tkh_ot , b.lokasi , b.nama  , b.gaji_pokok , round((b.gaji_pokok/3),2) as sepertiga" 
  ss = ss & " from payroll.pekerja_ot a , payroll.paymas b "
  ss = ss & "  where to_char(a.tkh_ot,'mm/yyyy') ='"& tkh_ot&"'  "
  ss = ss & " and a.no_pekerja = b.no_pekerja and b.lokasi = '"&ptj&"' " 
  ss =ss & " order by tkh_ot " 
  Set rsss = objConn.Execute (ss)
  tkh_ot = rsss ("tkh_ot")
  pemohon = rsss ("no_pekerja")	
  lokasi = rsss ("lokasi")
  nama1 = rsss ("nama")
  gaji_pokok = rsss ("gaji_pokok")
  sepertiga = rsss ("sepertiga")
	
	
	
   m = " select kod , initcap(keterangan) keterangan from iabs.jabatan "
   m = m & " where kod = '"&ptj&"' "
   Set rsm = objConn.Execute(m)
   ket = rsm ("keterangan")			
%>
<br>
<br>


<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd" height="20%">
  <TBODY>
    <tr class="hd"> 
    <td colspan="8">Senarai Pekerja Melebihi Sepertiga <br>
    Jabatan <%=ket%><br>  Bulan <%=tkh_ot%></td>
    </tr> 
    <tr style="BACKGROUND-COLOR: #ffffff"><td colspan="7">&nbsp;</td></tr>
    <tr class="hd"> 
    <td width="2%"   height="17"  align="center">Bil</td>
    <td width="7%"   align="center">Bulan/Tahun</td>
    <td width="7%"   align="center">Jabatan</td>
    <td width="8%"   align="center">No Pekerja</td>
    <td width="40%"  align="center">Nama Pekerja</td>
    <td width="19%"  align="center">Anggaran OT (RM)</td>
    <td width="17%"  align="center">Tuntutan OT Sebenar (RM)</td>
    </tr>
  
   
<%  
bil=0
cetak = ""
Do While Not rsss.eof 
tkh_ot = rsss ("tkh_ot")
pemohon = rsss ("no_pekerja")
sepertiga = rsss ("sepertiga")
color="#EFF5F5"

   'semak amaun ot sebulan , senarai semua yang buat ot by bulan --> nadia (01082016)
   se = " select  nvl(sum(angg_ot_S),0) + nvl(sum(angg_ot_M),0) as total_ot_sebulan , nvl(sum(ot_sebenar_m),0) +  nvl(sum(ot_sebenar_s),0) as otsebenar  "
   se = se & " from payroll.pekerja_ot"
   se = se & " where to_char(tkh_ot,'mm/yyyy') ='"& tkh_ot &"' and no_pekerja='"&pemohon&"'" 
   Set rse= objConn.Execute (se)
   if not rse.bof and rse.eof then
   total_ot_sebulan = rse("total_ot_sebulan")
   otsebenar = rse("otsebenar")
   end if
   
   
  if cdbl(rse("total_ot_sebulan")) > cdbl(rsss("sepertiga")) then
	color = "yellow"
	cetak = "y"
	bil=bil + 1
	semuaxgst1 = semuaxgst1 +  cdbl(rse("total_ot_sebulan"))
	sebenar1 = sebenar1 + cdbl(rse("otsebenar"))
	else
	cetak = "x"
  end if
  
  
   sse = " select sum(nvl(angg_ot_s,0) + nvl(angg_ot_m,0)) anggaran , count(distinct no_pekerja) as kira , sum(nvl(ot_sebenar_m,0) +  nvl(ot_sebenar_s,0)) sebenar "
   sse = sse & " from payroll.pekerja_ot"
   sse = sse & " where to_char(tkh_ot,'mm/yyyy') ='"& tkh_ot &"' and no_pekerja='"&pemohon&"' " 
   Set rsse= objConn.Execute (sse)
   Do While Not rsse.eof 
   anggaran = rsse("anggaran")
   kira = rsse("kira")
   sebenar = rsse("sebenar")
   
%>


<form name="test<%=bil%>"  method="post" action="sp27c.asp" target="_new">
  <% if cetak="y" then%>
  <tr bgcolor="<%=color%>">
    <td rowspan="2" align="center"><%=bil%></td>
    <td rowspan="2" align="center"><%=rsss("tkh_ot")%></td>
    <td rowspan="2" align="center"><%=rsss("lokasi")%></td>
    <td rowspan="2" align="center"><%=rsss("no_pekerja")%></td>
    <td rowspan="2" align="left"><%=rsss("nama")%></td>
    <td rowspan="2" align="right"><%=rsse("anggaran")%></td>
    <td rowspan="2" align="right"><%=rsse("sebenar")%></td> 
  </tr>
  <tr bgcolor="#EFF5F5">
  </tr>
 <% end if %> 
 </form> 

<% 
rsse.movenext
loop 
rsss.movenext
loop 

'-----
rsss.close

'-----

%>

<tr bgcolor="#EFF5F5">
<td colspan="5" align="right"><B>&nbsp; Jumlah Keseluruhan (RM)&nbsp;</B></td>
<td width="17%" align="right"><b><%=formatnumber(semuaxgst1,2)%></b></td>
<td width="17%" align="right"><b><%=formatnumber(sebenar1,2)%></b></td>
</tr>
</tbody> 
</table>
<br>


  <TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
    <tr align="center">
    <td colspan="2">
   <form name="test<%=bil%>"  method="post" action="sp27c1.asp?tkh_ot=<%=tkh_ot%>&amenu='sp27c1'" target="_new">
   <input type="submit" name="b7" value="Cetak Laporan" class="button" onFocus="nextfield='done';">
    </form>
    </td>
     </tr>
  </table>


<%
objConn.close
 end sub %>
<input type="hidden" name="bilrec" value="<%=ctrz%>" >
</body>
</html>