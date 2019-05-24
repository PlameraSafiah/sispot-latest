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
blnthn = request.form("blnthn")
p7 = request.form("b7")	'cetak permohonan
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


  'check jabatan by id pekerja
  pekd = request.cookies("gnop")
  a =     "select jabatan from majlis.pengguna where no_pekerja = '"& pekd &"' "
  Set rsa = objConn.Execute(a)


  idd = "select a.nama, a.lokasi, b.keterangan, to_char(sysdate,'dd/mm/yyyy') tkh  from payroll.paymas a, iabs.jabatan b "
  idd = idd & " where a.no_pekerja = '"&pekd&"' and b.kod = a.lokasi "
  set idd2 = objConn.Execute(idd)
  'response.write idd
  lokasi = idd2 ("lokasi")
 
 
  q2="select distinct to_char(a.tkh_ot,'mm/yyyy') as tkh_ot , b.ptj , b.pengesahan "
  q2 = q2 & " from payroll.jadual_ot a , payroll.arahan_kerja b where a.id_ak=b.id_ak and b.ptj = '"&lokasi&"' and b.pengesahan='Y' order by tkh_ot asc"  
  set oq2 = objConn.Execute(q2)
  if not oq2.bof and oq2.eof then
  'response.write q2
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
           %>
      </select></td></tr>
  
  
<tr align="center"><td colspan="2"><input type="submit" name="proses" value="Hantar"></td></tr>
</table>
</form>
<%
end sub %>



<%
sub borang()
blnthn=request.form("blnthn")
ptj=plokasi
p7 = request.form("b7")	'cetak permohonan
tkh_ot=request.form("tkh_ot")


  'list semua pekerja yang membuat ot(bulanan)
  ss = " select distinct (a.no_pekerja) ,  to_char(a.tkh_ot,'mm/yyyy') as tkh_ot , b.lokasi , b.nama  , b.gaji_pokok , round((b.gaji_pokok/3),2) as sepertiga   " 
  ss = ss & " from payroll.pekerja_ot a , payroll.paymas b "
  ss = ss & "  where to_char(a.tkh_ot,'mm/yyyy') ='"& tkh_ot&"'  "
  ss = ss & " and a.no_pekerja = b.no_pekerja and b.lokasi = '"&ptj&"' " 
  ss =ss & " order by tkh_ot " 
  'response.write ss
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
    <tr style="BACKGROUND-COLOR: #ffffff"><td colspan="8">&nbsp;</td></tr>
    <tr class="hd"> 
    <td width="4%"   height="17"  align="center">Bil</td>
    <td width="14%"   align="center">Bulan/Tahun</td>
        <td width="10%"  align="center">Jabatan</td>
    <td width="13%"   align="center">No Pekerja</td>
    <td width="59%"  align="center">Nama Pekerja</td>
    </tr>
  
   
<%  
bil=0
cetak = ""
Do While Not rsss.eof 
tkh_ot = rsss ("tkh_ot")
pemohon = rsss ("no_pekerja")
sepertiga = rsss ("sepertiga")
color="#EFF5F5"

   'semak amaun ot sebulan
   se = " select  nvl(sum(angg_ot_S),0) + nvl(sum(angg_ot_M),0) as total_ot_sebulan "
   se = se & " from payroll.pekerja_ot"
   se = se & " where to_char(tkh_ot,'mm/yyyy') ='"& tkh_ot &"' and no_pekerja='"&pemohon&"'" 
   Set rse= objConn.Execute (se)
   'response.write se
   total_ot_sebulan = rse("total_ot_sebulan")
   
  if cdbl(rse("total_ot_sebulan")) > cdbl(rsss("sepertiga")) then
	color = "yellow"
	cetak = "y"
	bil=bil + 1
	else
	cetak = "x"
  end if
%>



<form name="test<%=bil%>"  method="post" action="sp27c.asp" target="_new">
  <% if cetak="y" then%>
  <tr bgcolor="<%=color%>">
    <td rowspan="2" align="center"><%=bil%></td>
    <td rowspan="2" align="center"><%=rsss("tkh_ot")%></td>
    <td rowspan="2" align="center"><%=rsss("lokasi")%></td>
    <td rowspan="2" align="center"><%=rsss("no_pekerja")%></td>
    <td rowspan="2" align="left"><%=rsss("nama")%></td>
   
  </tr>
  <tr bgcolor="#EFF5F5">
  </tr>
      <% end if %> 
 </form> 

<% 
rsss.movenext
loop 
%>
  
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


<% end sub %>
<input type="hidden" name="bilrec" value="<%=ctrz%>" >
</body>
</html>