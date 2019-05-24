<!--#include file="connection.asp" -->
<%response.buffer=true%>
<%response.cookies("amenu") = "sp16.asp"%>
<!--'#include file="spmenu.asp"-->


<html>
<head>
<title>Laporan Kelulusan Melebihi Sepertiga</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" href="css/menu1.css" REL="stylesheet">
<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%=color4%>">

<script type="text/javascript" src="menu1.js"></script>
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
function check(f){
	if (f.blnthn.value=="")
	{
		alert("Sila masukkan bulan dan tahun kelulusan");
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


<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%'=color4%>">

<% 
proses = request.form("proses")
bulan = request.form("bulan")
tahun = request.form("tahun")
ket = request.form("keterangan")
keterangan = request.form("keterangan")
ptj=plokasi
p7 = request.form("b7")	'cetak permohonan


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


    'set bulan/tahun - arahan kerja 
    q4="select distinct to_char(tkh_ot,'mm/yyyy') as tkh_ot"
    q4 = q4 & " from payroll.jadual_ot order by tkh_ot"  
    set oq4 = objConn.Execute(q4)
    tkh_ot2 = oq4 ("tkh_ot")

    
	'data pengguna login
    idd = "select a.nama, a.lokasi, initcap(b.keterangan) keterangan, to_char(sysdate,'dd/mm/yyyy') tkh  from payroll.paymas a, iabs.jabatan b "
    idd = idd & " where a.no_pekerja = '"&pekd&"' and b.kod = a.lokasi "
    set idd2 = objConn.Execute(idd)
    'response.write idd
    lokasi = idd2 ("lokasi")
	ket = idd2 ("keterangan")
   

   'semak bulan 
   q2="select distinct a.bulan , b.lokasi "
   q2 = q2 & " from payroll.selenggara_kerja a , payroll.paymas b where b.lokasi = '"& lokasi &"' and a.id = b.no_pekerja "
   q2 = q2 & " order by bulan asc"  
   set oq2 = objConn.Execute(q2)
   'response.write q2
   if not oq2.bof and oq2.eof then
   bulan = oq2 ("bulan")
   end if


   'semak tahun
   q2b="select distinct tahun "
   q2b = q2b & " from payroll.selenggara_kerja where bulan like '"& bulan & "' "
   q2b = q2b & " order by tahun asc"
   set oq2b = objConn.Execute(q2b)
   'response.write  q2b
   if not oq2.bof and oq2.eof then
   tahun = oq2b ("tahun")
   end if
   
%>
<br>


<br>
<form name="test" method="post" action="sp16.asp" onSubmit="return check(this)">
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
   <tr class="hd">
      <td colspan="2">Senarai Pekerja Memperolehi Kelulusan Tuntutan Melebihi Sepertiga</td>
    </tr>
    
  <tr>
  <td width="15%" align="right"><b>Bulan Kelulusan : &nbsp;</b></td>
  <td><!--<input type="text" name="blnthn" size="2" value="<%=bulan%>">-->
     <select size="1" name="bulan" onChange="submitForm1(this)">
          <option value="">Pilih Bulan</option>
          <% do while not oq2.EOF  %>
          <option <%if cstr(bulan) = cstr(oq2("bulan")) then%> selected <%end if%>value="<%=oq2("bulan")%>"><%=oq2("bulan")%></option>
          <%
            oq2.MoveNext
            loop
           %>
    </select>
    </td>
    </tr> 
    
 
  <tr>
  <td align="right"><b>Tahun Kelulusan : &nbsp;</b></td>
  <td><!--<input type="text" name="pekerja" size="4" value="<%=tahun%>">-->
  <select size="1" name="tahun" onChange="submitForm1(this)">
  <option value="">Pilih Tahun</option>
          <% do while not oq2b.EOF %>
          <option <%if cstr(tahun) =  cstr(oq2b("tahun")) then%> selected <%end if%>value="<%=oq2b("tahun")%>"><%=oq2b("tahun")%></option>
          <%
            oq2b.MoveNext
            loop
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
ptj=plokasi
p7 = request.form("b7")	
bulan = request.form("bulan")
tahun = request.form("tahun")


 
   se = " select distinct a.id , b.nama"
   se = se & " from payroll.selenggara_kerja a , payroll.paymas b"
   se = se & " where a.id = b.no_pekerja and bulan = '"&bulan&"'  and tahun = '"&tahun&"' and lokasi='"& ptj &"' " 
   Set rse= objConn.Execute (se)
   'response.write se
   id = rse("id")
   nama = rse("nama")
  
  
  ff = " select keterangan from iabs.jabatan where kod='"&ptj&"' "
  Set rsff= objConn.Execute (ff)
  keterangan = rsff ("keterangan")
%>
<br>
<br>




<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all width="80%" align="center"> 
  <TBODY>
    <tr class="hd"> 
      <td colspan="10">Senarai Pekerja Yang Mendapat Kelulusan Melebihi Sepertiga<br> 
      Jabatan : <%=ptj%>-<%=keterangan%><br> 
      Bulan : <%=bulan%>/<%=tahun%></td>
    </tr> 
    
    
    <tr style="BACKGROUND-COLOR: #ffffff"><td colspan="2">&nbsp;</td>
    </tr>
    
    
    <tr class="hd">  
    <td width="4%"   align="center">Bil</td>
    <td width="96%"  align="center">Senarai Nama Pekerja Yang Mendapat Kelulusan Tuntutan Melebihi Sepertiga</td>
    </tr>
 
<%  
bil=0
Do While Not rse.eof 
bil=bil + 1
%>

<form name="test<%=bil%>"  method="post" action="sp16.asp" target="_new">
  <tr bgcolor="#EFF5F5">
    <td align="center"><%=bil%></td>
    <td align="left"><%=rse ("id")%>-<%=rse ("nama")%></td>
  </tr>
 </form> 
  
<%  
rse.movenext
loop 
%>  
</table>
<br>
  

  <TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
    <tr align="center">
    <td colspan="2">
   <form name="test<%=bil%>"  method="post" action="sp16a.asp?bulan=<%=bulan%>&tahun=<%=tahun%>&ptj=<%=ptj%>&amenu='sp16a'" target="_new">
   <input type="submit" name="b7" value="Cetak Laporan" class="button" onFocus="nextfield='done';">
    </form>
    </td>
     </tr>
  </table>
	   
<% end sub %>
<input type="hidden" name="bilrec" value="<%=ctrz%>" >
</body>
</html>