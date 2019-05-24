<!--#include file="connection.asp" -->

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
function checkCarian(f)
{
		if (f.carian.value=="")
	{
		alert("Sila masukkan no kakitangan/nama");
		f.carian.focus();
		return false
	}
}
function check(f){
	if (f.id_unit.value=="")
	{
		alert("Sila pilih unit kumpulan");
		f.id_unit.focus();
		return false
	}	
	if (f.no_pekerja.value=="")
	{
		alert("Sila masukkan no kakitangan");
		f.no_pekerja.focus();
		return false
	}
}
//  End -->
</script>

</head>
<body leftmargin="0" topmargin="0">

<% 
proses = request.form("proses")
'ptj=plokasi
ptj=109
id_unit=request.form("id_unit")
if proses="" or proses="Cari" then 
	papar
else
	no_pekerja = request.form("no_pekerja")
	
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
q01="delete from payroll.unit_kakitangan where no_pekerja = "& no_pekerja &" and ptj="& ptj &""
objConn.Execute(q01)

'check kakitangan ada di jabatan berkenaan
q0="select no_pekerja,nama from payroll.paymas where no_pekerja="& no_pekerja &" and lokasi="& ptj &""
set r0 = objConn.Execute(q0)
	if not r0.bof and not r0.eof then
		'get nama unit
		q02="select upper(nama_unit) as nama_unit from payroll.unit_sispot where ptj="& ptj &" and id_unit like '"& id_unit &"'"
		set oq2 = objConn.Execute(q02)
		if not oq2.bof and not oq2.eof then
			q2="Insert into payroll.unit_kakitangan (no_pekerja,nama,ptj,id_unit,nama_unit) values "
			q2 = q2 & "("& no_pekerja &",'"& replace(r0("nama"),"'","''") &"',"& ptj &",'"& id_unit &"','"& oq2("nama_unit") &"')"
			objConn.Execute(q2) 
			rekod="simpan"
			msg="Rekod unit kumpulan "& no_pekerja &" selesai ditambah"
		else
			rekod="tiada"
			msg="Unit tidak  wujud.  Sila daftar kakitangan sekali lagi"
		end if
	else
		rekod="tiada"
		msg="Kakitangan "& no_pekerja &" tiada di jabatan anda.<br>Proses tidak diteruskan"
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
%>
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd" style="font-family: Trebuchet MS; font-size: 10pt;">
  <TBODY>
  <form name="tt" method="post" action="sp12bv1.asp" onSubmit="return checkCarian(this)">
    <tr style="BACKGROUND-COLOR: #ffffff">
      <td colspan="2"><div align="right"><b>Nama&nbsp;:&nbsp;</b></div></td>
      <td colspan="3" valign="middle">&nbsp;<input type="text" name="carian" size="40" maxlength="50">&nbsp;&nbsp;<input type="submit" name="proses" value="Cari"></td>
    </tr>
   </form>
 </TBODY>
 </TABLE><br>
 <%
Dim iPageSize,iPageCount,iPageCurrent,iRecordsShown,adUseClient
Dim S
adUseClient = 3
iPageSize = 25

If Request.QueryString("page") = "" Then
	iPageCurrent = 1		
Else
	iPageCurrent = CInt(Request.QueryString("page"))
End If
	
id_unit=request.form("id_unit")
carian=request.form("carian")
q1="select rowid,id_unit,nama_unit,no_pekerja,nama from payroll.unit_kakitangan where ptj="& ptj &" "
if carian <> "" then
q1 = q1 & "and upper(nama) like upper('%"& replace(carian,"'","''") &"%') "
end if
q1 = q1 & "order by nama_unit,no_pekerja asc"
Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.CursorLocation = adUseClient
	
rs1.PageSize = iPageSize
rs1.CacheSize = iPageSize

rs1.CursorLocation = 3	'ubah paging (aisyah-25052015)
rs1.Open q1, objConn
		'rs1Open seld, objConn, adOpenStatic
		
iPageCount = rs1.PageCount 
	
if rs1.bof and rs1.eof then %>
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
  <TBODY>
    <tr style="BACKGROUND-COLOR: #ffffff"><td><div align="center"><b>Tiada Rekod</b></div></td></tr>
 </TBODY>
 </TABLE>
<%
'else
q0="select id_unit,upper(nama_unit) as nama_unit from payroll.unit_sispot where ptj="& ptj &" order by nama_unit"
set oq1=objConn.Execute(q0)

kira=rs1.recordcount

If iPageCurrent > iPageCount Then iPageCurrent = iPageCount

If iPageCurrent < 1 Then iPageCurrent = 1

bil=0
bilangan=Request.QueryString("bilangan")		
ms=Request.QueryString("ms")
		
If bilangan <>"" and ms="next" then
	bil = bilangan
End If
If bilangan <>"" and ms="pre" then
	bil = bilangan
End If

			If iPageCount <> 0 Then
				rs1.AbsolutePage = iPageCurrent 
				iRecordsShown = 0
				count = 0
				Do While iRecordsShown <iPageSize And Not rs1.eof 
					iRecordsShown = iRecordsShown + 1
					count = count + 1
					bil=bil + 1
					rs1.movenext
				loop
			end if	

			if kira > 0 then	
%>
<br>
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
  <TBODY>
    <tr class="hd">
      <td colspan="5"><div align="center">
      <br>
  <table width="70%" border=0 cellpadding=0 cellspacing=0 align="center" style="font-family: Trebuchet MS; font-size: 10pt;">
    <tr>
      <td align="left" width="41%"> Jumlah Rekod: <%=kira%> </td>
      <td align="right" width="17%"> 
	  <% If iPageCurrent <> 1 Then %> <a href="sp12bv1.asp?page=1&bilangan=0&ms=pre&carian=<%=carian%>&id_unit=<%=id_unit%>"> 
        <img name="firstrec" border="0" src="images\slide_beginning_02.gif" width="21" height="21"></a> 
        <% End if %> 
		<% If iPageCurrent <> 1 Then %> 
		<a href="sp12bv1.asp?page=<%= iPageCurrent - 1 %>&bilangan=<%=bil-count-iPageSize%>&ms=pre&carian=<%=carian%>&id_unit=<%=id_unit%>"> 
        <img name="previous" border="0" src="images\slide_previous_02.gif" width="21" height="21"></a> 
        <% End If %> </td>
      <td align="center" width="27%"> &nbsp;Halaman <%=iPageCurrent%>/<%if iPageCount=0 then %>1<% else %><%=iPageCount%><% end if %>&nbsp; 
      </td>
      <td align="right" width="15%"> <% If iPageCurrent < iPageCount Then	%> 
        <a href="sp12bv1.asp?page=<%= iPageCurrent + 1 %>&bilangan=<%=bil%>&ms=next&carian=<%=carian%>&id_unit=<%=id_unit%>"> 
        <img name="next" border="0" src="images/slide_next_02.gif" width="21" height="21"></a> 
        <% End If %> <% If iPageCurrent < iPageCount Then
	bil = (iPageCount - 1) * iPageSize  %> 
	<a href="sp12bv1.asp?page=<%=iPageCount %>&bilangan=<%=bil%>&ms=next&carian=<%=carian%>&id_unit=<%=id_unit%>"> 
        <img name="lastrec" border="0" src="images/slide_end_02.gif" width="21" height="21"></a> 
        <% End If %> </td>
</tr>
</table></div></td>
    </tr>
    <tr class="hd"> 
      <td align="center">Bil</td>
      <td align="center">Unit</td>
      <td align="center">No Kakitangan</td>
      <td align="center">Nama</td>
      <td align="center">Proses</td>
    </tr>
  <form name="test0" method="post" action="sp12bv1.asp" onSubmit="return check(this)">
    <tr valign="middle"> 
      <td>&nbsp; </td>
      <td>&nbsp;<select name="id_unit">
          <option value="">Sila Pilih</option>
          <%do while not oq1.eof %>
          <option <%if id_unit=oq1("id_unit") then %> selected <%end if%>value="<%=oq1("id_unit")%>"><%=oq1("nama_unit")%></option>
          <%oq1.movenext
	loop%>
        </select></td>
      <td><div align="center">
        <input type="text" name="no_pekerja" size="5" maxlength="5">
        &nbsp;&nbsp; 
        <!--
    <a href="javascript:;" onclick="winBRopen('cal_popup.asp?FormName=test0&FieldName=tkh_bayar_gaji&Date=<%=Date()%>&CurrentDate=<%=Date()%>','popup_cal','241','206','no','no')"><img src="images/icon_pickdate.png" class="DatePicker" alt="Pilih Tarikh" /></a>	-->
      </div></td>
      <td>&nbsp;</td>
      <td> <input type="submit" name="proses" value="Simpan"> </td>
  </form></tr>
  <%  
bil = 0
		ctrz = 0
	
		bilangan=Request.QueryString("bilangan")
		page = Request.QueryString("page")
		ms=Request.QueryString("ms")
	
		If bilangan <>"" and ms="next" then
			bil = bilangan
		End If
		If bilangan <>"" and ms="pre" then
			bil = bilangan
		End If
		
		If iPageCount <> 0 Then
			rs1.AbsolutePage = iPageCurrent
   			iRecordsShown = 0
			count = 0				
	    
		ctrz = 0
		Do while iRecordsShown <iPageSize and not rs1.EOF
bil=bil + 1
'get nama
z1="select nama from payroll.paymas where no_pekerja="& rs1("no_pekerja") &""
set oz1=objConn.Execute(z1)
%>
  <tr> 
    <td><%=bil%></td>
    <td>&nbsp;<%=rs1("nama_unit")%></td>
    <td><div align="center"><%=rs1("no_pekerja")%></div></td>
    <td>&nbsp;<%=oz1("nama")%></td>
    <form name="test_hapus<%=bil%>"  method="post" action="sp12bv1.asp">
      <input type="hidden" name="rowid" value="<%=rs1("rowid")%>">
      <input type="hidden" name="no_pekerja" value="<%=rs1("no_pekerja")%>">
      <td> <input type="submit" name="proses" value="Hapus" onClick="return confirm('Anda Pasti Hapus Rekod?')"> 
      </td>
    </form>
  </tr>
  <%iRecordsShown = iRecordsShown + 1
		count = count + 1
		rs1.MoveNext
		Loop	 %></tbody>
</table>
<% 
end if
end if
end if
end sub %>
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
no_pekerja=request("no_pekerja")
If a<>"" then
q1="Delete from payroll.unit_kakitangan where rowid like '"& a &"'"
objConn.Execute(q1)
papar()%>
<br>
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
  <TBODY> 
  <tr>
    <td align="center"> Rekod <%=no_pekerja%> telah dihapus.</td>
  </tr>
</tbody></table>
<%
end if
end if
end sub%>
</body>
</html>