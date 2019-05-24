<!--#include file="connection.asp" -->
<!--'#include file="spmenu.asp"-->
<%response.buffer=true%>
<%response.cookies("amenu") = "sp12c.asp"%>
<html>
<head>
<title>Hierarki Pengguna</title>
<META content="text/html; charset=iso-8859-1" http-equiv=Content-Type>
<link type="text/css" href="menu.css" REL="stylesheet">
<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%=color4%>">
</head>


<FORM name="test" action="sp12_hapus.asp" method="post" >
<% 





'no=2
id_unit  =request("id_unit")
penyelia2=request("penyelia2")
no       =request.QueryString("no")
response.write no


del = "delete from payroll.unit_sispot_penyelia where  no = '"& no &"' "
'del = "delete from payroll.unit_sispot_penyelia where rowid like '"& a &"' "
set rsdel = ObjConn.execute(del)
response.Write del
%>
<script >
	alert("Hapus selesai.")
</script> 


 
  
  </TABLE>
     </form>
</BODY></HTML>

 