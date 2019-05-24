<!--#include file="connection.asp" -->
<!--'#include file="spmenu.asp"--><head>
<meta http-equiv="Content-Language" content="en-us">
</head>

<%response.buffer=true%>

<%	

Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"

Set objConn1=Server.CreateObject("ADODB.Connection")
objConn1.Open "dsn=12c;uid=majlis;pwd=majlis;"

set rs4 = server.createobject("ADODB.Recordset")
rs4.open "select * from payroll.ptj where ( kod <> 100 and kod <> 110 and kod <> 217 and kod <> 237 and kod <> 227 and kod <> 207 and kod <> 999) order by kod  ",objConn1
%>



<table border="1" width="830" cellspacing="0" cellpadding="0" height="177">
	<tr>
		<td width="818" align="center" height="40" bgcolor="#DEDEDE" colspan="4">
		<b>PENYELIA YANG TELAH DIDAFTAR</b></td>
	</tr>

	<tr>
		<td width="818" align="center" height="10" bgcolor="#00DDDD" colspan="4">&nbsp;
		</td>
	</tr>

	<tr>
		<td width="818" align="center" height="40" bgcolor="#00DDDD" colspan="4">
		<form method="POST" action="sp18.asp">
			<p>Jabatan :&nbsp; <font color="#FFFFFF">
			<select size="1" name="jabatan" style="font-family: Calibri; font-size: 10pt">
			<option selected value="SILA PILIH">SILA PILIH</option>
			<%Do while not rs4.eof%>
			<option value="<%=rs4("kod")%>"><%=rs4("keterangan")%></option>
			<%rs4.movenext
			loop%>
			</select></font>&nbsp;&nbsp;
			<font color="#FFFFFF">
			<input type="submit" value="CARI" name="B1"></font></p>
		</td>
	</tr>

<%if request("jabatan") <> "SILA PILIH" and request("jabatan") <> "" then%>
	<tr>
		<td width="49" align="center" height="32" bgcolor="#BBBBBB">
		<font face="Calibri">Bil</font></td>
		<td width="333" align="center" height="32" bgcolor="#BBBBBB">
		<font face="Calibri">Nama</font></td>
		<td width="149" align="center" height="32" bgcolor="#BBBBBB">
		<font face="Calibri">No. Kakitangan</font></td>
		<td width="274" align="center" height="32" bgcolor="#BBBBBB">
		<font face="Calibri">Jabatan</font></td>
	</tr>

<%


set rs = server.createobject("ADODB.Recordset")

rs.open "select distinct no_pekerja from kebenaran_2002 where sistem = 'sp' AND skrin = 'sp31' ",objConn

b=0
nama=""
lokasi=""

Do while not rs.eof


set rs2 = server.createobject("ADODB.Recordset")
rs2.open "select * from payroll.paymas where no_pekerja = '"&rs("no_pekerja")&"' and lokasi = '"&request("jabatan")&"' ",objConn1

if not rs2.eof then
	b=b+1
	set rs3 = server.createobject("ADODB.Recordset")
	rs3.open "select * from payroll.ptj where kod = '"&rs2("lokasi")&"' ",objConn1

	nama=rs2("nama")
	lokasi=rs3("keterangan")

%>
	<tr>
		<td width="49" align="center">
		<font face="Calibri">&nbsp;<%=b%>.</font></td>
		<td width="333"><font face="Calibri">&nbsp;<%=nama%></font></td>
		<td width="149" align="center">
		<font face="Calibri">&nbsp;<%=rs("no_pekerja")%></font></td>
		<td width="274"><font face="Calibri">&nbsp;<%=lokasi%></font></td>
	</tr>
<%	end if

rs.movenext
loop

end if
%>
		</form>

</table>