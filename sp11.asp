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
	if (f.tkh_bayar_gaji.value=="")
	{
		alert("Sila masukkan tarikh pembayaran gaji");
		f.tkh_bayar_gaji.focus();
		return false
	}
}


//  End -->
</script>

</head>
<body leftmargin="0" topmargin="0">
<%response.cookies("amenu") = "sp11.asp"%>
<% 
proses = request.form("proses")
if proses="" then 
	papar
else
	tkh_bayar_gaji = request.form("tkh_bayar_gaji")
	
		if proses="Simpan" then 
			simpan
		elseif proses="Hapus" then 
			hapus
		end if
end if


sub simpan()
proses=request("proses")

if proses="Simpan" then
rekod=""
bln = mid(tkh_bayar_gaji,3,2)
thn = mid(tkh_bayar_gaji,5,4)
q0="select * from payroll.tkh_gaji where to_char(tkh_bayar_gaji,'mmyyyy') like '"& bln&thn &"'"
set r0=objConn.Execute(q0)
	if r0.eof then
		q2="Insert into payroll.tkh_gaji (tkh_bayar_gaji) values (to_date('"& tkh_bayar_gaji &" 23:59','ddmmyyyy hh24:mi'))"
		objConn.Execute(q2) 
		rekod="simpan"
		msg="Rekod tarikh akhir submit OT "& tkh_bayar_gaji &" selesai ditambah"
	else
		rekod="duplicate"
		msg="Rekod tarikh akhir submit tuntutan OT untuk bulan "& bln &" tahun "& thn &" telah wujud.<br>Proses tidak diteruskan"
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
end sub

sub papar()
q1="select rowid,to_char(tkh_bayar_gaji,'ddmmyyyy') as tkh_bayar_gaji,tkh_bayar_gaji as tkhsusun from payroll.tkh_gaji order by tkhsusun asc"
Set rs1 = objConn.Execute(q1)
%>
<br>
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
  <TBODY>
  <tr style="BACKGROUND-COLOR: #ffffff"><td colspan="3">&nbsp;</td></tr>
  <tr class="hd"> 
    <td align="center" width="10%">Bil</td>
    <td align="center">Tarikh (ddmmyyyy)</td>
    <td align="center" width="15%">Proses</td>
  </tr>
      <form name="test0" method="post" action="sp11.asp" onSubmit="return check(this)">
  <tr valign="middle"> 
    <td>&nbsp; </td>
    <td><input type="text" name="tkh_bayar_gaji" size="8" maxlength="8">&nbsp;&nbsp;
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
    <td><%=rs1("tkh_bayar_gaji")%></td>
    <form name="test_hapus<%=bil%>"  method="post" action="
    
    "> 
	<input type="hidden" name="rowid" value="<%=rs1("rowid")%>">
    <input type="hidden" name="tkh_bayar_gaji" value="<%=rs1("tkh_bayar_gaji")%>">    
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
a=request("rowid")
tkh_bayar_gaji=request("tkh_bayar_gaji")
If a<>"" then
q1="Delete from payroll.tkh_gaji where rowid like '"& a &"'"
objConn.Execute(q1)
papar()%>
<br>
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
  <TBODY> 
  <tr>
    <td align="center"> Rekod tarikh akhir submit tuntutan<%=tkh_bayar_gaji%> telah dihapus.</td>
  </tr>
</tbody></table>


<%
end if
end sub%>
<%sub edit()
rowid=request("rowid")
tkh_bayar_gaji=request("tkh_bayar_gaji")
tkh_bayar_gaji_asal=request("tkh_bayar_gaji_asal")
q1="Update payroll.tkh_gaji set tkh_bayar_gaji = to_date('"& tkh_bayar_gaji &"','ddmmyyyy') where rowid like '"& rowid &"'"
'objConn.Execute(q1)
papar()%>
<br>
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
  <TBODY> 
  <tr>
    <td align="center"> Rekod tarikh gaji <%=tkh_bayar_gaji_asal%> selesai diedit kepada <%=tkh_bayar_gaji%>.</td>
  </tr>
</tbody></table>
<%end sub%>
</body>
</html>