<!--#include file="connection.asp" -->
<%response.buffer=true%>
<%response.cookies("amenu") = "sp24.asp"%>
<!--'#include file="spmenu.asp"-->
<html>
<head>
<title>Sistem Pengurusan OT</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" href="css/menu1.css" REL="stylesheet">
<script type="text/javascript" src="menu1.js"></script>
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
function check(f){
	if (f.tkh_post_ak.value=="")
	{
		alert("Sila masukkan tarikh pengesahan arahan kerja");
		f.tkh_post_ak.focus();
		return false
	}
}


//  End -->
</script>

</head>
<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%'=color4%>">
<% 
'delete 0115110257 dr buku vot 07/01/2016
proses = request.form("proses")
tkh_post_ak=request.form("tkh_post_ak")
blnthn = request.form("blnthn")
pemohon = request.form("pemohon")
ptj=plokasi

nopekerja=request.cookies("gnop")

papar
if proses <> "" then	



		if proses="Hantar" then 
		     senarai
		elseif proses="Senarai Pekerja Mengikut Arahan Kerja" then
		     response.redirect"sp116.asp"
			 'response.redirect"sp116.asp"
		elseif proses="Sah" then 
			sah
			senarai
		end if
		
	
end if	
	



sub papar ()


mm = " select to_char (sysdate,'ddmmyyyy') as tkh_today from dual "
set mm3 = objConn.Execute(mm)

q1="select distinct a.pemohon as pemohon,upper(b.nama) as nama from payroll.arahan_kerja a,payroll.paymas b where a.pengesahan not like 'Y' and a.ptj="& ptj &" and b.no_pekerja=a.pemohon "
'q1="select distinct a.pemohon as pemohon,upper(b.nama) as nama from payroll.arahan_kerja a,payroll.paymas b where a.ptj="& ptj &" and b.no_pekerja=a.pemohon "
q1 = q1 & "order by a.pemohon asc"
set oq1 = objConn.Execute(q1)
%>
<br>
<form name="test" method="post" action="sp24.asp" onSubmit="return check(this)">
  <TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
    <tr> 
      <td><b>Pemohon</b></td>
      <td><select name="pemohon">
	  <option value="">Sila pilih</option>
       <% if nopekerja=13264 then%>
      <option value="16835">16835 - JAAFAR BIN MORAD (Unit Pelancongan ,Seni dan Warisan) </option>
      <%end if%>
          <% do while not oq1.eof %>
          <option <%if Cdbl(pemohon)=Cdbl(oq1("pemohon")) then%>selected <%end if%>value="<%=oq1("pemohon")%>"><%=oq1("pemohon")%>&nbsp;-&nbsp;<%=oq1("nama")%></option>
          <%oq1.movenext
		loop%>
        </select></td>
    </tr>
    <tr>
      <td><b>Bln/Thn Arahan Kerja</b>&nbsp;&nbsp;<font color="#FF0000"><b>'(mmyyyy)'</b></font></td>
      <td><input type="text" name="blnthn" size="5" maxlength="6" value="<%=blnthn%>"></td>
    </tr>
    <tr> 
      <td width="20%"><b> Tarikh Pengesahan</b>&nbsp;&nbsp;<font color="#FF0000"><b>'(ddmmyyyy)'</b></font></td>
      <td><%=mm3("tkh_today")%><input type="hidden" name="tkh_post_ak" value="<%=mm3("tkh_today")%>"></td>
    </tr>
    <tr align="center"> 
      <td colspan="2"><input type="submit" name="proses" value="Hantar">&nbsp;&nbsp;
       <!--<input type="submit" name="proses" value="Senarai Pekerja Mengikut Arahan Kerja">-->
       <a href="sp116.asp" target="_new">Senarai Pekerja Mengikut Arahan Kerja</a></td>
    </tr>
  </table>
</form>
<%
end sub %>
<%
sub senarai()
tkh_post_ak=request.form("tkh_post_ak")
if nopekerja=13264 then 'saf tambah 28032018
q1="select id_ak,keterangan,keterangan_1,nota,pemohon,bulan||'/'||tahun as bulantahun from payroll.arahan_kerja where tkh_sah is null "
else
q1="select id_ak,keterangan,keterangan_1,nota,pemohon,bulan||'/'||tahun as bulantahun from payroll.arahan_kerja where tkh_sah is null and ptj="& ptj &" "
end if
'q1="select id_ak,keterangan,nota,pemohon,bulan||'/'||tahun as bulantahun from payroll.arahan_kerja where tkh_sah is null and ptj="& ptj &" "
if pemohon <> "" then
q1 = q1 & "and pemohon="& pemohon &" "
end if
if blnthn <> "" then
q1 = q1 & "and tahun="& mid(blnthn,3,4) &" and bulan="& mid(blnthn,1,2) &" "
end if
q1 = q1 & "order by tahun,bulan,id_ak"
set oq1=objConn.Execute(q1)
%>
<br>
<br>
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
  <TBODY>
    <tr class="hd"> 
      <td colspan="10" >Senarai Arahan Kerja</td>
    </tr>
    <tr style="BACKGROUND-COLOR: #ffffff"> 
      <td colspan="10">&nbsp;</td>
    </tr>
    <tr class="hd"> 
      <td align="center">Bil</td>
      <td align="center">No Rujukan Arahan Kerja</td>
      <td align="center">Bln/Thn</td>
      <td align="center">Pemohon</td>
      <td align="center">Keterangan</td>
      <td align="center">Nota Tambahan</td>
      <td align="center">Bil Pekerja</td>
      <td align="center">Anggaran OT (RM) </td>
      <td align="center">Proses</td>
    </tr>
    <%  
bil=0
tot_angg_semua=0
Do While Not oq1.eof
bil=bil + 1
q2="select upper(nama) as nama from payroll.paymas where no_pekerja="& oq1("pemohon") &""
set oq2 = objConn.Execute(q2)
If not oq2.bof and not oq2.eof then
	nama = oq2("nama")
end if
q3="select count(distinct no_pekerja) as bil,nvl(sum(angg_ot_S),0) + nvl(sum(angg_ot_M),0) as tot_angg from payroll.pekerja_ot where id_ak='"& oq1("id_ak") &"'"
set oq3 = objConn.Execute(q3)
if not oq3.bof and not oq3.eof then
	tot_angg_semua = cdbl(tot_angg_semua) + cdbl(oq3("tot_angg"))
end if
%>
  <form name="test<%=bil%>"  method="post" action="sp24.asp">
    <input type="hidden" name="pemohon" value="<%=pemohon%>">
    <input type="hidden" name="blnthn" value="<%=blnthn%>">
    <tr> 
      <td><%=bil%></td>
      <td> <b><%=oq1("id_ak")%></b></td>
      <td ><%=oq1("bulantahun")%></td>
      <td ><%=oq2("nama")%>&nbsp;(<%=oq1("pemohon")%>)</td>
      <td><%=oq1("keterangan_1")%></td>
      <td><%=oq1("nota")%></td>
      <td><div align="center"><%=oq3("bil")%></div></td>
      <td><div align="right"><%=oq3("tot_angg")%></div></td>
      <td> <input type="hidden" name="tkh_post_ak" value="<%=tkh_post_ak%>"> <input type="hidden" name="id_ak" value="<%=oq1("id_ak")%>"> 
        <div align="center">
          <%if cdbl(oq3("bil")) > 0 then 
		  
		  pel = "select penyelia from payroll.arahan_kerja where penyelia = '"& nopekerja &"' "
		  Set Objpel = ObjConn.execute(pel)
		  'response.write pel
		  
		  if not Objpel.eof then
		  
		  nop = Objpel("penyelia")
		  
		  end if
 'response.write nop 
		  		  
		  if   nop = "" then
		  
		  keterangan = " Sila Rujuk Penyelia yang didaftar"
		  
		  
		  %> 
           
         <%=keterangan%>
        
         <% else %>
          
          <input type="submit" name="proses" value="Sah">

          <%end if 
		  end if%>
        </div></td>
    </tr>
  </form>
  <%
  oq1.movenext
  loop %>
  <tr> 
    <td colspan="7"><div align="right"><b>Jumlah Keseluruhan Anggaran OT</b></div></td>
    <td><div align="right"><%=formatnumber(tot_angg_semua,2)%></div></td>
    <td>&nbsp;</td>
  </tr></tbody>
</table>
<br>
<% end sub %>



<%
sub sah
id_ak=request.form("id_ak")
tkh_post_ak=request.form("tkh_post_ak")
u1="update payroll.arahan_kerja set pengesahan='Y',tkh_sah=to_date('"& tkh_post_ak &"','ddmmyyyy') where id_ak='"& id_ak &"'"
'response.Write(u1)
objConn.Execute(u1)

u1="update payroll.pekerja_ot set pengesahan='Y',tkh_sah=to_date('"& tkh_post_ak &"','ddmmyyyy') where id_ak='"& id_ak &"'" 
objConn.Execute(u1)

end sub
%>
</body>
</html>