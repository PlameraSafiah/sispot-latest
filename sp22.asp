<!--#include file="connection.asp" -->
<!--'#include file="spmenu.asp"-->
<%response.buffer=true%>
<%response.cookies("amenu") = "sp22.asp"%>
<html>
<head>
<title>Sistem Pengurusan OT</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" href="css/menu1.css" REL="stylesheet">
<script type="text/javascript" src="menu1.js"></script>
<script language="javascript">
function check1(f){
	if(f.id.value == ""){
	alert("Sila Pilih");
	return false;
	}
}

</script>


</head>
<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%=color4%>">
<% 
ptj=plokasi
no_pekerja=request.cookies("gnop")
proses = request.form("proses")


	fid = Request.Form("id")
	session("id") = fid
	
	
if proses = "Hantar" then	

 response.redirect "sp22a.asp?noid="&fid&""
 response.cookies("noid") = fid
end if




%>
<form name="test" method="post" action="sp22.asp" onSubmit="return check1(this)" >

<%
q1="select arahan_kerja.id_ak from "
q1 = q1 & "payroll.arahan_kerja where arahan_kerja.pemohon="& no_pekerja &" and "
q1 = q1 & "ptj="& ptj &"  order by tahun desc,bulan desc,id_ak"
set oq1 = objConn.Execute(q1)


%>

<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
<tr class="hd">
  <td colspan="2">Mengedit Arahan Kerja</td></tr>
<tr>
  <td width="15%"><b> No Rujukan Arahan Kerja (Belum post):</b> </td><td><select name="id">
  <%do while not oq1.eof%>
    <option <%if id=oq1("id_ak") then%>selected <%end if%>value="<%=oq1("id_ak")%>"><%=oq1("id_ak")%></option>
  <%oq1.movenext
  loop 
oq1.close
objConn.close
%>
  </select></td></tr>
<tr align="center"><td colspan="2"><input type="submit" name="proses" value="Hantar"></td></tr>
</table>
</form>
</body>
</html>