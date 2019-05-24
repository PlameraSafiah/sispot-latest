<!--#include file="connection.asp" -->
<!--'#include file="spmenu.asp"-->
<%response.buffer=true%>
<html>
<head>
<title>Sistem Pengurusan OT</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" href="css/menu1.css" REL="stylesheet">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" href="css/menu1.css" REL="stylesheet">
<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%=color4%>">
<script type="text/javascript" src="menu1.js"></script>
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
function check(f){
	if (f.nama_unit.value=="")
	{
		alert("Sila masukkan unit");
		f.nama_unit.focus();
		return false
	}
}

function upperCase(f) {
	if (f.nama_unit.value!="")
	{
		var tempStr = f.nama_unit.value.toUpperCase();
		f.nama_unit.value = tempStr;
	}
}
//  End -->
</script>

</head>
<body leftmargin="0" topmargin="0">
<%response.cookies("amenu") = "sp12a.asp"%>
<% 
proses = request.form("proses")
ptj=plokasi
if proses="" then 
	papar
else
	nama_unit = request.form("nama_unit")
	
		if proses="Simpan" then 
			simpan
		elseif proses="Hapus" then 
			hapus
		end if
end if


sub simpan()
if ptj="" then
%>
<script>
alert("Session anda telah tamat.  Sila login semula");
document.location.href= "http://mpspnet.mpsp.gov.my/sistemnet.asp";
</script>
<%
else
if proses="Simpan" then
rekod=""
q0="select * from payroll.unit_sispot where upper(nama_unit) like upper('"& nama_unit &"') and ptj="& ptj &""
set r0=objConn.Execute(q0)
	if r0.eof then
		a1="select (nvl(max(substr(id_unit,2,4)),0)) + 1 as siri from payroll.unit_sispot  where "
		a1 = a1 & "upper(substr(nama_unit,1,1)) like upper(substr('"& nama_unit &"',1,1))"
		set qa = objConn.Execute(a1)
		'siri = Cint(qa("siri") + 1
		
		d1 = " select substr( upper('"& nama_unit &"'),1,1)||lpad('"&qa("siri")&"',4,'0') as unit_id from dual "
        Set od1 = objConn.Execute(d1)
		
		q2="Insert into payroll.unit_sispot (id_unit,nama_unit,ptj) values ('"& od1("unit_id") &"',upper('"& nama_unit &"'),"& ptj &")"
		objConn.Execute(q2) 
		rekod="simpan"
		msg="Rekod unit "& nama_unit &" selesai ditambah"
	else
		rekod="duplicate"
		msg="Rekod "& nama_unit &" telah wujud.<br>Proses tidak diteruskan"
	end if
	papar
	if rekod<> "" then %><br>
    <TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
      <TBODY> 
      <tr>
        <td align="center"><%=msg%></td>
      </tr>
    </tbody></table>
<% end if
end if
end if
end sub

sub papar()
q1="select rowid,nama_unit as nama_unit,id_unit from payroll.unit_sispot where ptj="& ptj &" order by nama_unit asc"
Set rs1 = objConn.Execute(q1)
%>
<br>
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
  <TBODY>
  <tr style="BACKGROUND-COLOR: #ffffff"><td colspan="3">&nbsp;</td></tr>
  <tr class="hd"> 
    <td align="center">Bil</td>
    <td align="center">Unit</td>
    <td align="center">Proses</td>
  </tr>
      <form name="test0" method="post" action="sp12a.asp" onSubmit="return check(this)">
  <tr valign="middle"> 
    <td>&nbsp; </td>
    <td><input type="text" name="nama_unit" size="49" maxlength="49"  onBlur="return upperCase(this)">&nbsp;&nbsp;
    <!--
    <a href="javascript:;" onclick="winBRopen('cal_popup.asp?FormName=test0&FieldName=tkh_bayar_gaji&Date=<%=Date()%>&CurrentDate=<%=Date()%>','popup_cal','241','206','no','no')"><img src="images/icon_pickdate.png" class="DatePicker" alt="Pilih Tarikh" /></a>	-->
    </td>
    <td> 
        <input type="submit" name="proses" value="Simpan">
      </td>
    </form>
  </tr>
  <%  
bil=0
Do While Not rs1.eof 
bil=bil + 1
%>
  <tr> 
  	<td><%=bil%></td>
    <td><%=rs1("nama_unit")%></td>
    <form name="test_hapus<%=bil%>"  method="post" action="sp12a.asp"> 
	<input type="hidden" name="rowid" value="<%=rs1("rowid")%>">
    <input type="hidden" name="nama_unit" value="<%=rs1("nama_unit")%>"> 
	<input type="hidden" name="id_unit" value="<%=rs1("id_unit")%>">   
    <td>
      <input type="submit" name="proses" value="Hapus" onClick="return confirm('Anda Pasti Hapus Rekod?')">
    </td>
    </form>
  </tr>
  <%rs1.movenext
  loop %>
  </tbody> 
</table>
<% end sub %>
<%sub hapus()
if ptj="" then
%>
<script>
alert("Session anda telah tamat.  Sila login semula");
document.location.href= "http://mpspnet.mpsp.gov.my/sistemnet.asp";
</script>
<%
else
a=request("rowid")
b = request.form("id_unit")
nama_unit=request("nama_unit")
If a<>"" then
q1="Delete from payroll.unit_sispot where rowid like '"& a &"'"
objConn.Execute(q1)
q1="Delete from payroll.unit_kakitangan where id_unit like '"& b&"'"
objConn.Execute(q1)
papar()%>
<br>
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
  <TBODY> 
  <tr>
    <td align="center"> Rekod <%=nama_unit%> telah dihapus.</td>
  </tr>
</tbody></table>


<%
end if
end if
end sub%>
</body>
</html>