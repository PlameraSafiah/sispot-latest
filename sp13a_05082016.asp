<!--#include file="connection.asp" -->
<%response.buffer=true%>
<!--'#include file="spmenu.asp"-->
<%response.cookies("amenu") = "sp13a.asp"%>



<html>
<head>
<title>Kemaskini Kelulusan Sepertiga</title>
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
<script type="text/javascript" src="http://code.jquery.com/jquery-latest.min.js"/></script>
<script type="text/javascript" src="../sispot/javascript/jquery-latest.min.js"/></script>
<script type="text/javascript">
$(document).ready(function(){
$('input[name="all"],input[name="title"]').bind('click', function(){
var status = $(this).is(':checked');
$('input[type="checkbox"]', $(this).parent('li')).attr('checked', status);
});

});
</script>
<script type="text/javascript" src="menu1.js"></script>
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
function check(f){
	if (f.id.value=="")
	{
		alert("Sila masukkan no pekerja");
		f.id.focus();
		return false
	}
	if (f.bulan.value=="")
	{
		alert("Sila masukkan bulan");
		f.bulan.focus();
		return false
	}
	if (f.tahun.value=="")
	{
		alert("Sila masukkan tahun");
		f.tahun.focus();
		return false
	}
	
}
//  End -->
</script>
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
ptj=plokasi
ket = request.form("keterangan")
keterangan = request.form("keterangan")
p7 = request.form("b7")	'cetak permohonan
kod = request.form("kod")

	papar
	if proses="Simpan" then 
	  simpan
	elseif proses="Hapus" then 
      hapus
	elseif proses <> "" then
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

    pekd = request.cookies("gnop")
    a =     "select jabatan from majlis.pengguna where no_pekerja = '"& pekd &"' "
    Set rsa = objConn.Execute(a)
    
    'senarai jadual ot 
    q4="select distinct to_char(tkh_ot,'mm/yyyy') as tkh_ot"
    q4 = q4 & " from payroll.jadual_ot order by tkh_ot"  
    set oq4 = objConn.Execute(q4)
    tkh_ot2 = oq4 ("tkh_ot")
    
    'details id login
    idd = "select a.nama, a.lokasi, initcap(b.keterangan) keterangan, to_char(sysdate,'dd/mm/yyyy') tkh  from payroll.paymas a, iabs.jabatan b "
    idd = idd & " where a.no_pekerja = '"&pekd&"' and b.kod = a.lokasi "
    set idd2 = objConn.Execute(idd)
    lokasi = idd2 ("lokasi")
	ket = idd2 ("keterangan")

   
   'semak bulan kelulusan ---> nadia (19042016)
   if pekd=9678 or pekd=11208 then
   q2="select distinct bulan  "   
   q2 = q2 & " from payroll.selenggara_kerja"
   q2 = q2 & " order by bulan asc"  
   else
   q2="select distinct a.bulan , b.lokasi "   
   q2 = q2 & " from payroll.selenggara_kerja a , payroll.paymas b where b.lokasi = '"& lokasi &"' and a.id = b.no_pekerja "
   q2 = q2 & " order by bulan asc"  
   end if
   set oq2 = objConn.Execute(q2)
   'response.write q2
   if not oq2.bof and oq2.eof then
   bulan = oq2 ("bulan")
   end if
   

   'semak tahun kelulusan ---> nadia (19042016)
   if pekd=9678 or pekd=11208 then
   q2b="select distinct tahun   "
   q2b = q2b & " from payroll.selenggara_kerja where bulan like '"& bulan & "' "
   q2b = q2b & " order by tahun asc"
   else
   q2b="select distinct a.tahun , b.lokasi  "
   q2b = q2b & " from payroll.selenggara_kerja a , payroll.paymas b where b.lokasi  = '"& lokasi &"' and a.id = b.no_pekerja and bulan like '"& bulan & "' "
   q2b = q2b & " order by tahun asc"
   end if
   set oq2b = objConn.Execute(q2b)
   'response.write q2b
   if not oq2.bof and oq2.eof then
   tahun = oq2b ("tahun")
   end if



    'semak jabatan
	if pekd=9678 or pekd=11208 then
    qq2="select distinct b.lokasi , c.keterangan"
    qq2 = qq2 & " from payroll.selenggara_kerja a, payroll.paymas b , iabs.jabatan c where a.id = b.no_pekerja and b.lokasi = c.kod "
    qq2 = qq2 & " and bulan like '"& bulan & "' and tahun like '"& tahun & "' order by b.lokasi asc" 
	else 
	qq2="select distinct b.lokasi , c.keterangan"
    qq2 = qq2 & " from payroll.selenggara_kerja a, payroll.paymas b , iabs.jabatan c where a.id = b.no_pekerja and b.lokasi = c.kod "
    qq2 = qq2 & " and bulan like '"& bulan & "' and tahun like '"& tahun & "' and b.lokasi  = '"& lokasi &"' order by b.lokasi asc" 
	end if
    set oqq2 = objConn.Execute(qq2)
	'response.write qq2
	if not oq2.bof and oq2.eof then
    kod = oqq2 ("lokasi")
	keterangan = oqq2 ("keterangan")
	end if

   
%>
<br>


<br>
<form name="test" method="post" action="sp13a.asp" onSubmit="return check(this)">
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
   <tr class="hd">
      <td colspan="2">Semak &amp; Kemaskini Data Kelulusan Sepertiga</td>
    </tr>
    
   <tr>
   <td width="15%" align="right"><b>Bulan Kelulusan : &nbsp;</b></td>
   <td><!--<input type="text" name="blnthn" size="2" value="<%=bulan%>">-->
     <select size="1" name="bulan" onChange="submitForm1(this)">
          <option value="">Pilih Bulan</option>
          <% do while not oq2.EOF  %>
          <option <%if cdbl(bulan) = cdbl(oq2("bulan")) then%> selected <%end if%>value="<%=oq2("bulan")%>"><%=oq2("bulan")%></option>
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


  <tr>
  <td width="15%" align="right"><b>Jabatan  : &nbsp;</b></td>
  <td><!--<input type="text" name="blnthn" size="8" value="<%=kod%>">-->
     <select size="1" name="kod">
          <option value="">Pilih Jabatan</option>
          <% do while not oqq2.EOF %>
          <option <%if cstr(kod) = cstr(oqq2("lokasi")) then%> selected <%end if%>value="<%=oqq2("lokasi")%>"><%=oqq2("lokasi")%>-<%=oqq2("keterangan")%></option>
          <%
            oqq2.MoveNext
            loop
           %>
    </select>
    </td>
    </tr> 

<tr align="center"><td colspan="2"><input type="submit" name="proses" value="Hantar"><br>
<b><a href="../sispot/sp13.asp" target="_new">Kembali</a></b></font></td></tr>
</table>
</form>
<% end sub %>


<%
sub borang()
ptj=plokasi
p7 = request.form("b7")	
bulan = request.form("bulan")
tahun = request.form("tahun")
id = request.form("id")
kod = request.form("kod")
 
   'check jabatan by id login
    pekd = request.cookies("gnop")
    a =     "select jabatan from majlis.pengguna where no_pekerja = '"& pekd &"' "
    Set rsa = objConn.Execute(a)
    kod4 = rsa ("jabatan")

    'senarai jadual ot
    q4="select distinct to_char(tkh_ot,'mm/yyyy') as tkh_ot"
    q4 = q4 & " from payroll.jadual_ot order by tkh_ot"  
    set oq4 = objConn.Execute(q4)
    tkh_ot2 = oq4 ("tkh_ot")
  
    
	'data pengguna login
    idd = "select a.nama, a.lokasi, b.keterangan , to_char(sysdate,'dd/mm/yyyy') tkh  from payroll.paymas a, iabs.jabatan b "
    idd = idd & " where a.no_pekerja = '"&pekd&"' and b.kod = a.lokasi "
    set idd2 = objConn.Execute(idd)
    lokasi = idd2 ("lokasi")
	ket = idd2 ("keterangan")
	

   if kod <> "" then 
   ff = " select keterangan , kod from iabs.jabatan where kod='"&kod&"' "
   Set rsff= objConn.Execute(ff)
   keterangan = rsff("keterangan")
   kodjab = rsff ("kod")
   else
   ff = " select keterangan , kod from iabs.jabatan where kod='"&kod&"' "
   Set rsff= objConn.Execute(ff)
   if not rsff.bof and rsff.eof then
   keterangan = rsff("keterangan")
   kodjab = rsff ("kod")
   end if
   end if


   'papar senarai data selenggara sepertiga = nadia (19042016)
   if pekd <> 9678 or pekd <> 11208 then
   'q1="select a.rowid ,a.id as id , a.tahun as tahun, a.bulan as bulan from payroll.selenggara_kerja a , payroll.paymas b where a.id=b.no_pekerja "
     q1="select distinct a.id as id , a.tahun as tahun, a.bulan as bulan from payroll.selenggara_kerja a , payroll.paymas b where a.id=b.no_pekerja "
   q1 = q1 & " and b.lokasi = '"&kod&"'  and a.bulan ='"&bulan&"' and a.tahun='"&tahun&"' order by a.tahun , a.bulan , a.id  asc"
   Set rs1 = objConn.Execute(q1)
   'response.write q1
   if not rs1.bof and rs1.eof then
    id = rs1 ("id")	
   end if	
   else
   'q1="select a.rowid ,a.id as id , a.tahun as tahun, a.bulan as bulan from payroll.selenggara_kerja a , payroll.paymas b where a.id=b.no_pekerja "
     q1="select distinct a.id as id , a.tahun as tahun, a.bulan as bulan from payroll.selenggara_kerja a , payroll.paymas b where a.id=b.no_pekerja "
   q1 = q1 & " and b.lokasi = '"&kod&"' and a.bulan ='"&bulan&"' and a.tahun='"&tahun&"' order by a.tahun , a.bulan , a.id  asc"
   Set rs1 = objConn.Execute(q1)
   'response.write q1
   if not rs1.bof and rs1.eof then
    id = rs1 ("id")	
   end if
   end if
  

%>
<br>
<br>


<TABLE border=1 borderColor=black cellPadding=0 cellSpacing=0 rules=all class="hd">
  <TBODY>    
    <tr>
  	<td align="center" class="hd" colspan="6"><B>SENARAI KELULUSAN TUNTUTAN MELEBIHI SEPERTIGA<br></B></td>
    <td width="0%"></td>
  </tr>
  
  
  <tr>
  	<td width="2%"   align="center" class="hd"><B>Bil</B></td>
    <td width="25%"  align="center" class="hd"><B>Nama Pekerja</B></td>
    <td width="20%"  align="center" class="hd"><B>Jabatan</B></td>
    <td width="13%"  align="center" class="hd"><B>Bulan</B></td>
    <td width="11%"  align="center" class="hd"><B>Tahun</B></td>   
    <td width="29%"  align="center" class="hd"><B>Proses Selenggara Data</B>
    </td>
  </tr>
  </tbody> 
</table>


<%  
bil=0
Do While Not rs1.eof 
bil=bil + 1
id = rs1 ("id")


mm = "select nama , lokasi from payroll.paymas where no_pekerja='"& id & "' "
Set rsmm = objConn.Execute(mm)
nama = rsmm ("nama")
lokasi = rsmm ("lokasi")


nn = "select keterangan from iabs.jabatan where kod='"& lokasi & "' "
Set rsnn = objConn.Execute (nn)
ket = rsnn ("keterangan")

%>



<TABLE border=1 borderColor=black cellPadding=0 cellSpacing=0 rules=all class="hd">
  <TBODY> 
  <tr>
  	<td width="2%"  align="center"><%=bil%></td>
    <td width="25%" align="left">&nbsp;<%=rs1("id")%>-<%=rsmm("nama")%></td>
    <td width="20%" align="left">&nbsp;(<%=rsmm("lokasi")%>) - <%=ket%></td>
    <td width="13%" align="center"><%=rs1("bulan")%></td>
    <td width="11%" align="center"><%=rs1("tahun")%></td>
    <form name="test_hapus<%=bil%>"  method="post" action="sp13a.asp"> 
    <input type="hidden"  name="id"      value="<%=rs1("id")%><">
    <input type="hidden"  name="bulan"   value="<%=rs1("bulan")%>">
    <input type="hidden"  name="tahun"   value="<%=rs1("tahun")%>">    
    <td width="29%" align="center"><input type="submit" name="proses" value="Hapus" onClick="return confirm('Anda Pasti Hapus Rekod?')">
    </td>
    </form>
  </tr>
  
  <% 
  rs1.movenext
  loop 
  %>

  </tbody> 
</table>
<% end sub %>


<%sub hapus()
a=request("rowid")
id=request("id")
bulan=request("bulan")
tahun=request("tahun")


If a<>"" then
q1="Delete from payroll.selenggara_kerja where rowid like '"& a &"' "
objConn.Execute(q1)
'response.write q1
%>
<br>
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
  <TBODY> 
  <tr>
    <td align="center">Rekod <b>><%=id%></b> bagi bulan <b><%=bulan%> , <%=tahun%></b> ini telah dihapus</td>
  </tr>
</tbody>
</table>
<%end if%>
<%end sub%>

</body>
</html>