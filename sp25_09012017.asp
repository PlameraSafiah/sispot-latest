<!--#include file="connection.asp" -->
<%response.buffer=true%>
<%response.cookies("amenu") = "sp25.asp"%>
<!--'#include file="spmenu.asp"-->
<html>
<head>
<title>Sistem Pengurusan OT</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" href="css/menu1.css" REL="stylesheet">
<script type="text/javascript" src="menu1.js"></script>
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
function check1(f){
	if(f.blnthn.value==""){
      alert("Sila Masukkan Bulan/Tahun")
      f.blnthn.focus();
	  return false;
     }
	 if(f.blnthn.value != ""){
		var valid="0123456789";
		var ok = "yes";
		var temp;
		for (var i=0; i<f.blnthn.value.length; i++) {
		temp = "" + f.blnthn.value.substring(i, i+1);
		if (valid.indexOf(temp) == "-1") ok = "no";
		}
		if (ok == "no") {
		alert("Kemasukan Data Tidak Diterima! Sila Masukkan Nombor Sahaja");
		f.blnthn.focus();
		f.blnthn.select();
		return false;
   		}
		else if(ok == "yes"){
		if((f.blnthn.value.length > 6) ||(f.blnthn.value.length < 6)){
		alert("Pastikan 6 digit dimasukkan.\nie:mmyyyy")
		f.blnthn.focus();
		return false;
		}
		var nilai1=f.blnthn.value.substring(0,2)
		var nilai2=f.blnthn.value.substring(2,6)
		if ((nilai2 < 1) || (nilai2 > 9999)){
		alert ("Tahun tidak sah");
		f.blnthn.focus();
		return false;
		}
		else if ((nilai1 > 12) || (nilai1 < 1)){
		alert("Bulan tidak sah");
		f.blnthn.focus();
		return false;
		}
		} 
	 } 	
}
//  End -->
</script>

</head>
<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%'=color4%>">
<% 
proses = request.form("proses")
blnthn = request.form("blnthn")
ptj=plokasi
papar
if proses <> "" then	
		if proses="Hantar" then 
			senarai
		elseif proses="Batal" then 
			batal
			senarai
		end if
end if

sub papar ()
if blnthn = "" then
Dim TodayYearNumber, TodayMonthName
TodayYearNumber = DatePart("yyyy",Date)
TodayMonth = DatePart("m",Date)
if len(TodayMonth)="1" then TodayMonth="0"& TodayMonth
blnthn=TodayMonth&TodayYearNumber
end if
%>
<br>
<form name="test" method="post" action="sp25.asp" onSubmit="return check(this)">
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
<tr><td width="15%"><b>Bulan/Tahun(mmyyyy):</b> </td><td><input type="text" name="blnthn" size="5" maxlength="6" value="<%=blnthn%>"></td></tr>
<tr align="center"><td colspan="2"><input type="submit" name="proses" value="Hantar"></td></tr>
</table>
</form>
<%
end sub %>
<%
sub senarai()
q1="select id_ak,keterangan,nota,pemohon,to_char(tkh_sah,'dd/mm/yyyy') tkh_sah from payroll.arahan_kerja where "
q1 = q1 & "tahun="& mid(blnthn,3,4) &" and bulan="& mid(blnthn,1,2) &" and tkh_sah is not null and ptj="& ptj &" "
q1 = q1 & "order by tahun desc,bulan desc,id_ak"
set oq1=objConn.Execute(q1)
%>
<br>
<br>
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
  <TBODY>
    <tr class="hd"> 
      <td colspan="10" >Senarai Arahan Kerja Yang Telah Dibuat Pengesahan</td>
    </tr>
    <tr style="BACKGROUND-COLOR: #ffffff"> 
      <td colspan="10">&nbsp;</td>
    </tr>
    <tr class="hd"> 
      <td align="center">Bil</td>
      <td align="center">No Rujukan Arahan Kerja</td>
      <td align="center">Tkh Pengesahan</td>
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
sah=""
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

q3c="select count(no_pekerja) as bil from payroll.pekerja_ot where id_ak='"& oq1("id_ak") &"' and status_claim is not null"
set oq3c = objConn.Execute(q3c)
if oq3c("bil") > 0 then sah="ada"
%>
  <form name="test<%=bil%>"  method="post" action="sp25.asp">
  <input type="hidden" name="blnthn" value="<%=blnthn%>">
    <tr> 
      <td><%=bil%>&nbsp;</td>
      <td><%=oq1("id_ak")%>&nbsp;</td>
      <td ><div align="center"><%=oq1("tkh_sah")%></div></td>
      <td ><%=oq2("nama")%>&nbsp;(<%=oq1("pemohon")%>)</td>
      <td><%=oq1("keterangan")%>&nbsp;</td>
      <td><%=oq1("nota")%>&nbsp;</td>
      <td><div align="center"><%=oq3("bil")%></div></td>
      <td><div align="right"><%=formatnumber(oq3("tot_angg"),2)%></div></td>
      <td><input type="hidden" name="id_ak" value="<%=oq1("id_ak")%>"> 
        <div align="center">
        <!---- buang checking dulu sementara nadia (10112016) ----->
        <%if sah="ada" then%>
        <!---Tuntutan Arahan Kerja Sudah Dibuat.<br>Pengesahan arahan kerja tidak boleh dibatalkan.--->
        <input type="submit" name="proses" value="Batal">
        <%else%>
        <input type="submit" name="proses" value="Batal">
        <%end if%>
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
<% end sub 

sub batal
id_ak=request.form("id_ak")
tkh_post_ak=request.form("tkh_post_ak")
u1="update payroll.arahan_kerja set pengesahan='T',tkh_sah=null where id_ak='"& id_ak &"'"
objConn.Execute(u1)

u1="update payroll.pekerja_ot set pengesahan='T',tkh_sah=null where id_ak='"& id_ak &"'" 
objConn.Execute(u1)
end sub
%>
</body>
</html>