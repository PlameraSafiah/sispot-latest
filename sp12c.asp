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
<script language="javascript">
 function data_simpan(a)
 { alert(a+" Data Di Simpan !!! ");
   return(true);
 }
 function data_update(a)
 { alert(a+" Data Di Ubah !!! ");
   return(true);
 }
</script>


<FORM name="kew" action="sp12c.asp" method="post" >
 
<%
  Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"


ptj=plokasi 
'response.Write(ptj)
  b = request.form("b")
  p = Request.Form("B1")
  d = request.form("d")
  
  proses = Request.Form("proses") 
  penyelia = request.form("penyelia")
  pelulus = request.form("pelulus")
  fnama=request.form("fnama")




if b = "Tambah " then 	
  	response.redirect "penyelia.asp"
end if

 if p = "Simpan" then
			sv = "insert into payroll.unit_sispot_pelulus "
		sv = sv & " (no_pekerja,lokasi) "
		sv = sv & " values ('"& fnama &"', '"& ptj &"') "
		
		set svr = objConn.execute(sv)
      'response.write sv
	  
	  response.write "<script language = ""vbscript"">"
	  response.write " MsgBox ""Data Pelulus Ditambah!"", vbInformation, ""Perhatian!"" "
	  response.write "</script>"
	  
end if   

 bilrec = Request.form("bilrec")
	rec = Request.form("rec")
	if rec <> "" then
			for i=1 to rec
				b1 = "b1" + cstr(i)
				's1 = "s1" + cstr(i)
				rowid = "frowid" + cstr(i)
				penyelia = "penyelia" + cstr(i)
				pelulus = "pelulus" + cstr(i)
                                id_unit = "id_unit" + cstr(i) 'saf 27032018--> 
				fbil = "fbil" + cstr(i)
				fnama = "fnama" + cstr(i)
				
				b1 = Request.form(""& b1 &"")
				fbil = Request.form(""& fbil &"")
				fnama = Request.form(""& fnama &"")
				's1 = Request.form(""& s1 &"")
				penyelia = Request.form(""&penyelia &"")
				pelulus = Request.form(""&pelulus &"")
                                id_unit = Request.form(""&id_unit &"") 'saf 27032018--> 
				
	if b1 = "Hapus" then
			dt = "delete from payroll.unit_sispot_penyelia where nama_penyelia ='"& penyelia&"' and id_unit='"&id_unit&"'"
			Set objRs3 = objConn.Execute(dt)

	  response.write "<script language = ""vbscript"">"
	  response.write " MsgBox ""Data Di Padam!"", vbInformation, ""Perhatian!"" "
	  response.write "</script>"
	  
	       end if	
		   
		 if b1 = "Hapus1" then
		' pelulus = Request.form("pelulus")
			dt1 = "delete from payroll.unit_sispot_pelulus where no_pekerja ='"&pelulus&"' "
			Set rsdt1 = objConn.Execute(dt1)
           ' response.write dt1
	  response.write "<script language = ""vbscript"">"
	  response.write " MsgBox ""Data Pelulus Di Padam!"", vbInformation, ""Perhatian!"" "
	  response.write "</script>"
	  
	       end if   
		    
			next			
		end if
		
		
		
   	 	  

    ' -------------------------------nama jabatan utk tajuk---------------------------------------------------
    b0 = "select keterangan,kod from iabs.jabatan where kod = '"& ptj &"'  "
    Set rsb0 = objConn.Execute(b0)
	if not rsb0.eof then 
	jabatan=rsb0("keterangan")
	end if
	
	
  
%>  
<p align="center">&nbsp;</p>
<p align="center"><strong>SENARAI PELULUS MENGIKUT JABATAN <%=jabatan %></strong></p>
<table width="90%"  border="0" align="center" >
  
  
 
   <TR align="center" bgcolor="<%=color1%>"> 
    <TD width="7%" align="center" valign="middle">Bil</TD>
    <TD  width="75%"><b>NAMA PELULUS</b></TD>
    <td width="18%">Proses</td>
  </TR>
   <TR align="center" bgcolor="<%=color1%>"> 
    <TD width="7%"  valign="middle"></TD>
    <TD align="left" width="75%"><select size="1" name="fnama" >

            <%	g = " select no_pekerja, UPPER(nama) as nama from payroll.paymas where lokasi = '"&ptj&"'  order by nama "
    	Set objRsg = objConn.Execute(g) 
		
		
			if fnama <> "" then  		
  		gg = " select no_pekerja, UPPER(nama) as nama from payroll.paymas where no_pekerja = '"& fnama &"' and  lokasi = '"&ptj&"' order by nama "
    	Set objRsgg = objConn.Execute(gg)    
			response
    	Do while not objRsgg.eof   %>
              <option value="<%=objRsgg("no_pekerja")%>" ><%=objRsgg("no_pekerja")%> - <%=objRsgg("nama")%></option>
              <%	objRsgg.MoveNext
     	            loop
		else    			%>
              <option value=""> Pilih Pelulus</option>
              <%	end if
    	do while not objRsg.EOF 		%>
              <option value="<%=objRsg("no_pekerja")%>"><%=objRsg("no_pekerja")%> - <%=objRsg("nama")%></option>
              <%  	objRsg.MoveNext
			     	loop  %>
            </select></TD>
    <td><input type="submit"  name="b1" value="Simpan"  ></td>
  </TR>
 <!-- -----------------------papar pelulus----------------------------->
  <TR  bgcolor="<%=color3%>"> 
  <% q= "select a.no_pekerja,b.nama,b.no_pekerja from payroll.unit_sispot_pelulus a,payroll.paymas b where a.lokasi='"& ptj &"' and a.no_pekerja=b.no_pekerja  "
    Set rsq = objConn.Execute(q)
	'response.Write( q)
	ctrz=0
	do while not rsq.eof
	'no = cdbl(no)+1
	ctrz = cdbl(ctrz) + 1
	pelulus=rsq("no_pekerja")
	nama=rsq("nama")
   %>
   
   
    <TD bgcolor="<%=color2%>" width="7%" align="center" valign="middle"><%=ctrz%></TD>
    <TD bgcolor="<%=color2%>" width="75%"><b><%=pelulus%>-<%=nama%></b>
      <input type="hidden" name="pelulus<%=ctrz%>" value="<%=pelulus%>" ></TD>
    <td bgcolor="<%=color2%>" align="center"> 
    <input type="submit" onClick="return confirm('Hapus Rekod? Pelulus = <%=pelulus%>')" name="b1<%=ctrz%>" value="Hapus1"  ></td>
 </TR>
<%rsq.movenext
  loop %> 
</table>
  
<!--  '--------------------------------------------------------penyelia------------------------------------------------------------->
<p align="center">&nbsp;</p>
<p align="center"><strong>SENARAI PENYELIA JABATAN <%=jabatan %> MENGIKUT UNIT</strong></p>
<table width="90%"  border="0" align="center" >
  
  <tr bgcolor="<%=color1%>"> 
  <td colspan="4"  align="center" style="color:#F00"><strong> * Untuk PENAMBAHAN PENYELIA, sila click Tambah.</strong></td> </tr>
  <tr bgcolor="<%=color1%>">      
    <td colspan="4"  align="center" width="32%" > <input type="submit" name="b" value="Tambah " >
  </tr>
  
   <!--'************************************************-senarai yang dah ada-*****************************************************************************-->
    <TR align="center" bgcolor="<%=color1%>"> 
    <TD width="5%" align="center" valign="middle">Bil</TD>
    <TD  width="35%"><b>Unit</b></TD>
    <TD><b>Penyelia </b></TD>
    <td>Proses</td>
  </TR>
  
<%	
a= "select a.rowid,a.id_unit,a.no_penyelia,a.nama_penyelia,a.ptj,a.no,upper(b.nama) as nama from payroll.unit_sispot_penyelia a,payroll.paymas b where a.nama_penyelia=b.no_pekerja order by a.id_unit "
    Set rsa = objConn.Execute(a)

 'response.write a
set rsa = objconn.Execute(a)
ctrz = 0
bil = 0
	
	
Do while not Rsa.EOF
    bil = bil + 1
   	ctrz = cdbl(ctrz) + 1
    penyelia   =rsa("nama_penyelia")
 	no_penyelia=rsa("no_penyelia")
 	id_unit    =rsa("id_unit")
	no          =rsa("no")
	nama   = rsa("nama")
	rowid = rsa("rowid")
	'warna = ctrz mod 2	
	


    '----------------------------nama unit ------------------------------------------------------
	b1 = "select id_unit,nama_unit from payroll.unit_sispot where id_unit = '"& id_unit &"'  "
    Set rsb1 = objConn.Execute(b1)
  	
	if not rsb1.eof then
	nama_unit = rsb1("nama_unit")
  	end if
	

%>
 
 <tr bgcolor="<%=color2%>" align="center">
   <td><%=ctrz%>&nbsp;</td>
   <td > <b><%=id_unit %> - <%=nama_unit%></b></td>
   <td align="left"> &nbsp;&nbsp;[ <%=penyelia%> ] - <%=nama%>
   <input type="hidden" name="penyelia<%=ctrz%>" value="<%=penyelia%>" >   
   <input type="hidden" name="id_unit<%=ctrz%>" value="<%=id_unit%>" >  <!--'saf 27032018-->   </td>
   
   
  <!-- <td> <input type="submit" name="e<'%=ctrz%>"  value="Edit">&nbsp;</td>-->                                        
   <td>
     
     <input type="submit" onClick="return confirm('Hapus Rekod? Bank = <%=penyelia%>')" name="b1<%=ctrz%>" value="Hapus"  >
  
   
   
   
   </td> 
 </tr>

<% 

  
rsa.MoveNext
	Loop	

%>
<input type="hidden" name="bilrec" value="<%=bil%>" >
<input type="hidden" name="rec" value="<%=bil%>" >

</tbody>
</table>


