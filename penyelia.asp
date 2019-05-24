<!--#include file="connection.asp" -->
<!--'#include file="spmenu.asp"-->
<%response.buffer=true%>
<%response.cookies("amenu") = "penyelia.asp"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<link type="text/css" href="menu.css" REL="stylesheet">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>
<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%=color4%>">
<script language="javascript">
function invalid_nsave(a)
    {  
       alert (b+" Data Di Simpan !!! ");
	
		return true;
		
    }

function invalid_nupdate(b)
    {  
       alert (b+" Data Di Kemaskini !!! ");
	
		return true;
		
    }

</script>


<% 
ptj=plokasi 
'response.Write(ptj)
menu = Request.cookies("kmenu")
kodz = ucase(Request.form("kodz"))

proses1 = request.Form("B2")

funit = request.Form("funit")
fjenis = request.Form("fjenis")
fnama = request.Form("fnama")
''fptj = request("ptj")

if proses1 = "Kembali" then 
response.Redirect "sp12c.asp"
end if

if proses1 = "Simpan" then 

'response.Write funit 
'response.Write " --- " & fjenis & " --- "
'response.Write fnama & " --- "


sr = "select distinct(id_unit) id_unit, nama_penyelia from payroll.unit_sispot_penyelia "
sr = sr & "where id_unit = '"& funit &"' and nama_penyelia = '"& fnama &"'  "
set objsr = ObjConn.execute(sr)
'response.Write sr

if not objsr.eof then


		uv = "update payroll.unit_sispot_penyelia set id_unit = '"& funit &"', nama_penyelia = '"& fnama &"', ptj = '"& ptj &"' "
		uv = uv & "where id_unit = '"& funit &"' and nama_penyelia = '"& fnama &"' "

		set objuv = ObjConn.execute(uv)
		   'response.Write uv 
		   response.write "<script language=""javascript"">"
           response.write "var timeID = setTimeout('invalid_nupdate(""  "");',1) "
           response.write "</script>"
		   response.Redirect("sp112.asp")

else 

		sv = "insert into payroll.unit_sispot_penyelia "
		sv = sv & " (id_unit,nama_penyelia, ptj) "
		sv = sv & " values ('"& funit &"', '"& fnama &"', '"& ptj &"') "

		set svr = objConn.execute(sv)

	'response.Write sv
	response.write "<script language=""javascript"">"
    response.write "var timeID = setTimeout('invalid_nsave(""  "");',1) "
    response.write "</script>"
	''response.redirect "sp112.asp"

      response.write "<script language = ""vbscript"">"
	  response.write " MsgBox ""Data Di Simpan!"", vbInformation, ""Perhatian!"" "
	  response.write "</script>"

end if 

end if 

	''--------------------untuk displaykan data even page load --------------------------
	if funit <> "" then funit = replace(funit,"'","`") end if
	if fjenis <> "" then fjenis = replace(fjenis,"'","`") end if
	if fnama <> "" then fnama = replace(fnama,"'","`") end if	

  ' -------------------------------nama jabatan utk tajuk---------------------------------------------------
    b0 = "select keterangan,kod from iabs.jabatan where kod = '"& ptj &"'  "
    Set rsb0 = objConn.Execute(b0)
	if not rsb0.eof then 
	jabatan=rsb0("keterangan")
	end if

%>


<FORM name="test" action="penyelia.asp" method="post" >

<p align="center"><strong>TAMBAH PENYELIA JABATAN <%=jabatan %> MENGIKUT UNIT</strong></p>
<TABLE align='center'border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all  width="59%">
  <TR > 
<td bgcolor="<%=color1%>"> Unit : </td>
<td bgcolor="<%=color1%>">

<select size="1" name="funit" >
<%		g = " select id_unit, UPPER(nama_unit) as nama from payroll.unit_sispot where ptj = '"& ptj &"' order by id_unit "
    	Set objRsg = objConn.Execute(g)
		
		
			if funit <> "" then  		
  		gg = " select id_unit, UPPER(nama_unit) as nama from payroll.unit_sispot where id_unit = '"& ptj &"' order by id_unit "
    	Set objRsgg = objConn.Execute(gg)    
			
    	Do while not objRsgg.eof   %>
              <option value="<%=objRsgg("id_unit")%>" ><%=objRsgg("id_unit")%> 
              - <%=objRsgg("nama")%></option>
              <%	objRsgg.MoveNext
     	loop
		else    			%>
              <option value=""> Pilih Unit</option>
              <%	end if
    	do while not objRsg.EOF 		%>
              <option value="<%=objRsg("id_unit")%>"><%=objRsg("id_unit")%> - <%=objRsg("nama")%></option>
              <%  	objRsg.MoveNext
     	loop  %>
            </select>
            
            </td>

</TR>

<tr>   
<td bgcolor="<%=color1%>">
Senarai Penyelia : 
</td>  
<td bgcolor="<%=color1%>">

<select size="1" name="fnama" >

            <%	g = " select no_pekerja, UPPER(nama) as nama from payroll.paymas where status_gaji='1' and  lokasi = '"& ptj & "' order by nama "
    	Set objRsg = objConn.Execute(g) 
		'response.write g
		
			if fnama <> "" then  		
  		gg = " select no_pekerja, UPPER(nama) as nama from payroll.paymas where no_pekerja = '"& fnama &"' and  lokasi = '"& ptj & "' and status_gaji='1' order by nama "
    	Set objRsgg = objConn.Execute(gg)    
			
    	Do while not objRsgg.eof   %>
              <option value="<%=objRsgg("no_pekerja")%>" ><%=objRsgg("no_pekerja")%> 
              - <%=objRsgg("nama")%></option>
              <%	objRsgg.MoveNext
     	loop
		else    			%>
              <option value=""> Pilih Penyelia</option>
              <%	end if
    	do while not objRsg.EOF 		%>
              <option value="<%=objRsg("no_pekerja")%>"><%=objRsg("no_pekerja")%> - <%=objRsg("nama")%></option>
              <%  	objRsg.MoveNext
			     	loop  %>
            </select>

</td>


</tr>


<tr>
          <td bgcolor="<%=color1%>" class="hd" colspan="2" align="center" height="50"> 
          <input type="submit" name="B2" value="Simpan"  class="button" > 
          <input type="submit" name="B2" value="Kembali"  class="button" > 
</td></tr>







</TABLE>
</FORM>
</body>
</html>
