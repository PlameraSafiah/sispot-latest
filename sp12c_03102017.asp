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


<FORM name="test" action="sp12c.asp" method="post" >
<% 

ptj=plokasi 
proses         = Request("Submit")
proses2        = Request("Submit1")
proses3        = Request("Submit2")
penyelia1edit   = request("penyelia1edit")
penyelia2edit   = request("penyelia2edit")
penyelia3edit   = request("penyelia3edit")
pelulus1edit    = request("pelulus1edit")
pelulus2edit    = request("pelulus2edit")

penyelia1	   = request("penyelia1")
penyelia2	   = request("penyelia2")
penyelia3	   = request("penyelia3")
pelulus1       = request("pelulus1")
pelulus2        = request("pelulus2")
id_unit        = request("id_unit")




if proses2="Edit" then'and penyelia1edit<>"" and  penyelia2edit<>"" and  penyelia3edit<>""   
pelulus1edit      = request("pelulus1edit")
pelulus2edit        = request("pelulus2edit")
penyelia1edit   = request("penyelia1edit")
penyelia2edit   = request("penyelia2edit")
penyelia3edit   = request("penyelia3edit")
id_unit  	    = request("id_unit")

	e= "Update payroll.ptj_pelulus set pelulus1='"&pelulus1edit&"',pelulus2='"&pelulus2edit&"' where kod='"&ptj&"' "
	Set Rse =objConn.Execute (e)
	
	
	d = "Update payroll.unit_sispot set penyelia1 = '"& penyelia1edit &"',penyelia2 = '"& penyelia2edit &"',penyelia3 = '"& penyelia3edit &"'"
	d = d & "where id_unit = '"& id_unit &"'  "
	Set Rs3 = objConn.Execute(d)
	'response.write d
	%>
	<script language="JavaScript">
	alert("Edit selesai.")
	</script>
	
<%end if%>
<%
	q= "select kod,pelulus1,pelulus2 from payroll.ptj_pelulus  where kod='"& ptj &"'  "
    Set rsq = objConn.Execute(q)
	'response.Write( q1)
	
 
	 bil=0

%>
<TABLE align='center'border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all  width="24%">
 
  <TR align="center" bgcolor="<%=color1%>"> 
  <TD  width="27%"><b>Pelulus 1</b></TD>
  <TD  width="36%"><input type="text" name="pelulus1edit" size="8" maxlength="8" value=<%=rsq("pelulus1")%>></TD>
 <TD width="37%" align="center" valign="middle"><input name="Submit1" type="submit" value="Edit"></TD>
 </TR>
  <TR align="center" bgcolor="<%=color1%>"> 
  <TD  width="27%"><b>Pelulus 2</b></TD>
  <TD  width="36%"><input type="text" name="pelulus2edit" size="8" maxlength="8" value=<%=rsq("pelulus2")%>></TD>
  <TD align="center" valign="middle"><input name="Submit1" type="submit" value="Edit"></TD>
   <input type="hidden" name="pelulus1" value="<%=rsq("pelulus1")%>">
   <input type="hidden" name="pelulus2" value="<%=rsq("pelulus2")%>">
 </TR>
</TABLE>

<br> 

<TABLE align='center'border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all  width="59%">
  <TR align="center" bgcolor="<%=color1%>"> 
    <TD width="5%"   align="center" valign="middle">Bil</TD>
    <TD  width="38%"><b>Unit</b></TD>
    <TD  width="13%"><b>Penyelia 1</b></TD>
    <TD  width="14%"><b>Penyelia 2</b></TD>
    <TD  width="14%"><b>Penyelia 3</b></TD>
   
    
    <TD  width="16%"><b>Proses</b></TD>
  </TR>
    
 <%
 
 
	q1= "select rowid,nama_unit as nama_unit,id_unit,penyelia1,penyelia2,penyelia3 from payroll.unit_sispot  where ptj='"& ptj &"'  "
    Set rsq1 = objConn.Execute(q1)
	'response.Write( q1)
	
 
	 bil=0
Do While Not rsq1.eof
 bil=bil+1
 penyelia1=rsq1("penyelia1")
 penyelia2=rsq1("penyelia2")
 penyelia3=rsq1("penyelia3")
 id_unit=rsq1("id_unit")
  %>
  
  
   <tr> 
  	<td align="center"><%=bil%></td>
    <td><%=rsq1("nama_unit")%></td>
    <td align="center"><input type="text" name="penyelia1edit" size="8" maxlength="8" value=<%=rsq1("penyelia1")%>></td>
    <TD align="center"><input type="text" name="penyelia2edit" size="8" maxlength="8" value=<%=rsq1("penyelia2")%>></TD>
    <td align="center"><input type="text" name="penyelia3edit" size="8" maxlength="8" value=<%=rsq1("penyelia3")%>></td>
    <TD align="center" valign="middle"><input name="Submit1" type="submit" value="Edit">
                                 &nbsp;<!--<input name="Submit2" type="submit"  value="Hapus">--></TD>
  </TR>
 
	<input type="hidden" name="rowid" value="<%=rsq1("rowid")%>">
    <input type="hidden" name="nama_unit" value="<%=rsq1("nama_unit")%>"> 
	<input type="hidden" name="id_unit" value="<%=rsq1("id_unit")%>"> 
    <input type="hidden" name="penyelia1" value="<%=rsq1("penyelia1")%>">
    <input type="hidden" name="penyelia2" value="<%=rsq1("penyelia2")%>">
    <input type="hidden" name="penyelia3" value="<%=rsq1("penyelia3")%>">   
    
  
   <tr align="center">
    
   </tr>
 
  
   <%
  rsq1.movenext
  loop %>

  
  </TABLE>
     </form>
</BODY></HTML>

 