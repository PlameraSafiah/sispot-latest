<!--#include file="connection.asp" -->
<!--'#include file="spmenu.asp"-->
<%response.buffer=true%>
<%response.cookies("amenu") = "sp23.asp"%>
<html>
<head>
<title>Sistem Pengurusan OT</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" href="css/menu1.css" REL="stylesheet">
<script type="text/javascript" src="menu1.js"></script>
<script language="javascript">
function submitForm1()
{ 
	document.test.submit();
}
</script>
</head>
<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%=color4%>">
<% 
proses = request.form("proses")
id_ak = request.form("id_ak")
tkh_ot = request.form("tkh_ot")
rowid_jadual_ot = request.form("rowid_jadual_ot")
ptj=plokasi
carianpemohon = request.cookies("gnop")
papar

tahun = DatePart("yyyy",Date)
'response.Write tahun
bulantahun = DatePart("m",Date) &"/"& DatePart("yyyy",Date)
'response.Write bulantahun
'borang
if proses <> "" then	
		if proses="Hantar" then 			
			borang
		elseif proses="Hapus" then 
			hapus
			borang
		else
			tambah
			borang
		end if
end if

sub tambah()
'response.Write("nak tambah data")
pekerja=request.form("pekerja")
kadar_M = request.form("kadar_M")
kadar_S = request.form("kadar_S")
jam_M = request.form("jam_M")
jam_S = request.form("jam_S")
minit_M = request.form("minit_M")
minit_S = request.form("minit_S")
masa_mula =  request.form("masa_mula")
masa_tamat =  request.form("masa_tamat")
post = "tidak"

tkh_input=mid(tkh_ot,1,2) &"/01/"& mid(tkh_ot,5,4)
TodayDate = DatePart("m",Date) &"/01/"& DatePart("yyyy",Date)

'check ptj dan status gaji
c0="select lokasi as ptj from payroll.paymas where no_pekerja="& pekerja &""
set cc0 = objConn.Execute(c0)
'response.write c0
'check tkh sah dan tkh lulus claim ot
v1="select tkh_lulus from payroll.arahan_kerja where id_ak='"& id_ak &"' and tkh_sah is not null and tkh_lulus is not null"
set ov1 = objConn.Execute(v1)
if not ov1.bof and not ov1.eof then
 post="ya"
end if

if not cc0.bof and not cc0.eof then
	if Cstr(cc0("ptj")) <> Cstr(ptj) then
	%>
	<script language="javascript">
	alert("Pekerja <%=pekerja%> bukan di jabatan tuan.");
	</script>
	<%

elseif post="ya" then
%>
	<script language="javascript">
	alert("Arahan kerja sudah disahkan.  Proses tambahan tidak dpt dibuat");
	</script>
	<%


end if
'-----
cc0.close
'------


'-----------------------------------------------------------------filter by peruntukan jabatan-----------------------------------------------------
'check total anggaran tuntutan ot ikot jabatan
v2="select sum(angg_ot_s)as siang ,sum(angg_ot_m)as malam,sum(angg_ot_s + angg_ot_m)as total from PAYROLL.PEKERJA_OT where ptj ='"& ptj &"' "
v2= v2 & " AND tkh_ot BETWEEN TO_DATE ('01012017', 'dd/mm/yyyy')AND TO_DATE ('31122017', 'dd/mm/yyyy') "
set ov2 = objConn.execute(v2)
'response.write v2
if not ov2.bof and not ov2.eof then
agg_tuntutan = ov2("total")

end if 
'-----
ov2.close
'------

'check peruntukan untuk jabatan per Year
 'mimi 17/10/2017
  v3 = " select belanjawan, "
  v3 = v3 & " nvl(sum(nvl(peruntukan,0) + nvl(tambahan,0) + nvl(viremen,0)),0) peruntukan1, "
  v3 = v3 & " nvl(sum(tanggungan),0) tanggungan,nvl(sum(bayaran),0) bayaran "  
  v3 = v3 & " from iabs.buku_vot "
  v3 = v3 & " where tahun = '"&tahun&"' and ptj = '"& ptj &"' and belanjawan = '14101'"
  v3 = v3 & " group by belanjawan order by belanjawan"
  set ov3 = objconn.execute(v3) 
  'response.write v3
  if not ov3.bof and not ov3.eof then
  jum_peruntuk=jum_peruntuk + ov3("peruntukan1")
  end if 

'-----
ov3.close
'------
  
  if jum_peruntuk < agg_tuntutan then 
  peruntukkan = "TAK OK"
  end if
  
'------------------------------------------------------------------------------------------------------------------------------------------------------
if peruntukkan = "TAK OK" then 
%>

<!--<script language="javascript">
	alert("Melebihi peruntukkan jabatan");
	</script>-->
<%
end if
	

cc0v="select 'x' no_pekerja from payroll.proses_ot where no_pekerja='"& pekerja &"' and to_char(tkh_ot,'mmyyyy')='"& mid(tkh_ot,1,2) & mid(tkh_ot,5,4) &"' and sah_individu='Y'"
set qcc0 = objConn.Execute(cc0v)
if not qcc0.bof and not qcc0.eof then
	%>
	<script language="javascript">
	alert("Kakitangan <%=pekerja%> sudah membuat tuntutan bagi bulan <%=mid(tkh_ot,1,2)%>, tahun <%=mid(tkh_ot,5,4)%>.");
	</script>
	<%
else
c1="select no_pekerja from payroll.pekerja_ot where no_pekerja='"& pekerja &"' and masa_mula_angg=to_date('"& masa_mula &"','mmddyyyy HH24:MI') "
c1 = c1 & "and masa_tamat_angg=to_date('"& masa_tamat &"','mmddyyyy HH24:MI') and tkh_ot = to_date('"&tkh_ot&"','mmddyyyy')"
c1 = c1 & "and id_ak = '"&id_ak&"' "
set cc1 = objConn.Execute(c1)
''response.write c1
if not cc1.bof and not cc1.eof then
	%>
	<script language="javascript">
	alert("Rekod pekerja <%=pekerja%> sudah wujud.");
	</script>
	<%
else
	jamtotal_M=Cdbl(jam_M) + Cdbl(minit_M)
	jamtotal_S=Cdbl(jam_S) + Cdbl(minit_S)

end if

'
'tt="select no_pekerja from payroll.paymas where no_pekerja="& pekerja &" "
'set rstt=objConn.execute (tt)
'response.write tt
'if not rstt.eof then
'pekerja1=rstt("no_pekerja")
'end if

''response.write(jamtotal_S)
''response.write("<br>"&jamtotal_M)
if Cdate(TodayDate) > Cdate(tkh_input) then
		qs2="select round(((amaun * 12/2504)*'"& kadar_S &"')*'"& jamtotal_S &"',2) as angg_ot_S,"
		qs2= qs2 & "round(((amaun * 12/2504)*'"& kadar_M &"')*'"& jamtotal_M &"',2) as angg_ot_M, "
		qs2 = qs2 & " 'Blh' AS status_gaji from payroll.paymas,payroll.paymas_txn where "
		qs2 = qs2 & "paymas_txn.bulan='"& mid(tkh_ot,1,2) &"' and paymas_txn.no_pekerja='"& pekerja &"' and paymas_txn.tahun='"& mid(tkh_ot,5,4) &"' and "
		qs2 = qs2 & "paymas_txn.no_pekerja=paymas.no_pekerja"
else
	
		qs2="select round(((gaji_pokok * 12/2504)*'"& kadar_S &"')*'"& jamtotal_S &"',2) as angg_ot_S, "
		qs2= qs2 & "round(((gaji_pokok * 12/2504)*'"& kadar_M &"')*'"& jamtotal_M &"',2) as angg_ot_M, "
		qs2 = qs2 & "decode(status_gaji,'4',' ditahan gaji','8',' sudah pencen','9',' sudah meninggal','Blh') as status_gaji "
		qs2 = qs2 &" from payroll.paymas where no_pekerja='"& pekerja &"' "
end if
		''response.write("<br><br>qs2---->"& qs2)
		set oqs2 = ObjConn.Execute(qs2)
		
		if not oqs2.bof and not oqs2.eof then
			if oqs2("status_gaji")="Blh" then
				
					i1="insert into payroll.pekerja_ot (id_ak,ptj,no_pekerja,tkh_ot,masa_mula_angg,masa_tamat_angg,angg_ot_S,"
					i1 = i1 & "angg_ot_M,masa_masuk,masa_balik,"
					i1 = i1 &"jam_S,minit_S,kategori_hari,kategori_masa,kadar_S,jam_M,minit_M,kadar_M,"
					i1 = i1 &"kelulusan,tkh_lulus,tkh_sah,pengesahan) values "
					i1 = i1 & "('"& id_ak &"','"& ptj &"','"& pekerja & "',to_date('"& tkh_ot &"','mmddyyyy'),"
					i1 = i1 & "to_date('"& masa_mula &"','mmddyyyy HH24:MI'),"
					i1 = i1 & "to_date('"& masa_tamat &"','mmddyyyy HH24:MI'),'"& oqs2("angg_ot_S") &"','"& oqs2("angg_ot_M") &"',null,"
					i1 = i1 & "null,null,null,null,null,null,null,null,null,null,null,null,null)"
					'response.write("<br>i1---->"& i1)
					ObjConn.Execute(i1)
			else
			%>
			<script language="javascript">
			alert("Proses tidak dpt diteruskan kerana pekerja <%=pekerja%><%=oqs2("status_gaji")%>.");
			</script>
			<%
			end if
		end if
end if
'end if 'qcc0 eof
end if 'bkn pekerja jabatan
'end if 'cco eof
end sub

sub hapus()
rowid_pekerja_ot = request.form("rowid_pekerja_ot")
pekerja=request.form("pekerja")
d1="delete from payroll.pekerja_ot where rowid like '"& rowid_pekerja_ot &"'"
objConn.Execute(d1)

d1="delete from payroll.proses_ot where id_ak='"& id_ak &"' and no_pekerja ="& pekerja &""
objConn.Execute(d1)
end sub

sub papar ()
'saf ubah 10/11/2017 buang tkh_sah is null
q1="select id_ak from payroll.arahan_kerja where ptj="& ptj &"  and tkh_lulus is null and pemohon="& carianpemohon  &" order by tahun desc,bulan desc,id_ak desc" 'and id_ak in "
'q1 = q1 & "(select id_ak from jadual_ot) order by tahun desc,bulan desc,id_ak"
set oq1 = objConn.Execute(q1)
%>
<br>
<form name="test" method="post" action="sp23.asp">
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
<tr class="hd">
  <td colspan="2">Assign Arahan Kerja Pada Seorang Pekerja</td></tr>
<tr>
  <td width="20%"><b> No Rujukan Arahan Kerja (Belum post):</b> </td>
  <td><select name="id_ak" onChange="submitForm1(this)">
  <option value="">Sila Pilih</option>
  <%do while not oq1.eof %>
    <option <%if id_ak=oq1("id_ak") then %> selected <%end if%>value="<%=oq1("id_ak")%>"><%=oq1("id_ak")%></option>
    <%oq1.movenext
	loop%>
  </select></td></tr>
  <%if id_ak <> "" then 
q2="select distinct to_char(tkh_ot,'dd/mm/yyyy') as tkh_ot, to_char(tkh_ot,'mmddyyyy')as tkh_ot_value "
q2 = q2 & " from payroll.jadual_ot where id_ak='"& id_ak &"' order by tkh_ot asc"  
set oq2 = objConn.Execute(q2)



  %>
<tr>
  <td><b>Tarikh OT</b></td>
  <td><select name="tkh_ot" onChange="submitForm1(this)">
  <option value="">Sila Pilih</option>
  <%
  'biltkh=0
  do while not oq2.eof 
  'biltkh = biltkh + 1
 

  %>
    <option <%if tkh_ot=oq2("tkh_ot_value") then %> selected <%end if%>value="<%=oq2("tkh_ot_value")%>"><%=oq2("tkh_ot")%></option>
    <%oq2.movenext
	loop
'-----
oq2.close
'------
%>
  </select></td>
</tr>
<%if tkh_ot <> "" then 'or biltkh=1 then 
q3="select rowid,to_char(masa_mula,'HH24:MI') as masa_mula,"
q3 = q3 & "decode(to_char(masa_tamat,'HH24:MI'),'00:00','24:00',to_char(masa_tamat,'HH24:MI')) as masa_tamat "
q3 = q3 & "from payroll.jadual_ot where id_ak='"& id_ak &"' and tkh_ot like to_date('"& tkh_ot &"','mmddyyyy') order by masa_mula , masa_tamat"
set oq3 = objConn.Execute(q3)
%>
<tr>
  <td><b>Masa Mula</b></td>
  <td><select name="rowid_jadual_ot">
    <%do while not oq3.eof %>
    <option value="<%=oq3("rowid")%>"><%=oq3("masa_mula")%>-<%=oq3("masa_tamat")%></option>
    <%oq3.movenext
	loop
'-----
oq3.close
'------
%>
  </select></td>
</tr>
<tr align="center"><td colspan="2"><input type="submit" name="proses" value="Hantar"></td></tr>
<%end if%>
<%end if%>
</table>
</form>
<%
end sub %>



<%
sub borang()
'tkh_ot=request.form("tkh_ot")
ptj=plokasi
q3="select penyelia,pemohon,pengesahan,unit,ptj,keterangan,keterangan_1,nota,bulan,tahun from payroll.arahan_kerja where id_ak ='"& id_ak &"'"
set oq3 = objConn.Execute(q3)
if not oq3.eof then
	penyelia = oq3("penyelia")
	ptj = oq3("ptj")
	no_pekerja = oq3("pemohon")
	
	q2="select upper(nama) as nama from payroll.paymas where no_pekerja='"& no_pekerja &"'"
	set oq2 = objConn.Execute(q2)
	If not oq2.bof and not oq2.eof then
		pemohon = oq2("nama")
	end if
	q2="select upper(nama) as nama from payroll.paymas where no_pekerja='"& penyelia &"'"
	set oq2 = objConn.Execute(q2)
	If not oq2.bof and not oq2.eof then
		penyelia_nama = oq2("nama")
	end if
'-----
oq2.close
'------
end if

q6="select to_char(tkh_ot,'dd/mm/yyyy') as tkh_ot,to_char(masa_mula,'HH24:MI') as masa_mula,kadar_M,kadar_S,"
q6 = q6 & "decode(to_char(masa_tamat,'HH24:MI'),'00:00','24:00',to_char(masa_tamat,'HH24:MI')) as masa_tamat_display,"
q6 = q6 & "to_char(masa_tamat,'HH24:MI') as masa_tamat,jam_S,jam_M,minit_S,minit_M from payroll.jadual_ot where rowid like '"& rowid_jadual_ot &"'"
set oq6 = objConn.Execute(q6)
%>
<br>
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
  <TBODY>
  <tr>
    <td class="hd">No Rujukan Arahan Kerja</td>
    <td>&nbsp;<%=id_ak%></td>
  </tr>
  <tr><td class="hd">Bln&nbsp;/Thn</td>
  <td><%=oq3("bulan")%>/<%=oq3("tahun")%></td></tr>
  <tr>
    <td class="hd">Penyelia</td>
    <td><%=penyelia_nama%> (<%=penyelia%>)</td>
  </tr>
  <tr><td class="hd">Pemohon</td>
  <td><%=pemohon%> (<%=no_pekerja%>)</td></tr>
  <tr><td class="hd">Jabatan / Unit</td>
  <td><%=ptj%>&nbsp;/&nbsp;<%=oq3("unit")%></td></tr>
  <tr>
    <td class="hd">Tarikh</td>
    <td><%=oq6("tkh_ot")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="hd">Masa Mula (HH24:MM) </td>
  <td><%=oq6("masa_mula")%>&nbsp;</td></tr>
    <tr>
      <td class="hd">Masa Tamat (HH24:MM) </td>
    <td><%=oq6("masa_tamat_display")%>&nbsp;</td></tr>
    <tr>
      <td class="hd">Keterangan Kerja</td>
       <td><%if oq3("keterangan_1")<>"" then %>
	   <%=oq3("keterangan_1")%>
       <%else%>
	   <%=oq3("keterangan")%>&nbsp;
       <%end if%></td>
    </tr>
    <tr><td class="hd"> Nota Tambahan</td>
  <td><%=oq3("nota")%>&nbsp;</td></tr>  
  </TBODY>
  </TABLE>
<br>
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
  <TBODY>
    <tr class="hd"> 
      <td colspan="6" >Daftar Kakitangan</td>
    </tr>
  <form name="test0" method="post" action="sp23.asp">
    <input type="hidden" name="id_ak" value="<%=id_ak%>">
    <input type="hidden" name="tkh_ot" value="<%=tkh_ot%>">
    <input type="hidden" name="masa_mula" value="<%=tkh_ot%>&nbsp;<%=oq6("masa_mula")%>">
    <input type="hidden" name="masa_tamat" value="<%=tkh_ot%>&nbsp;<%=oq6("masa_tamat")%>">
    <input type="hidden" name="rowid_jadual_ot" value="<%=rowid_jadual_ot%>">
    <input type="hidden" name="kadar_M" value="<%=oq6("kadar_M")%>">
    <input type="hidden" name="kadar_S" value="<%=oq6("kadar_S")%>">
    <input type="hidden" name="jam_S" value="<%=oq6("jam_S")%>">
    <input type="hidden" name="minit_S" value="<%=oq6("minit_S")%>">
    <input type="hidden" name="jam_M" value="<%=oq6("jam_M")%>">
    <input type="hidden" name="minit_M" value="<%=oq6("minit_M")%>">
    <tr bgcolor="#EFF5F5"> 
      <td>No Pekerja</td>
      <td colspan="4">&nbsp;
        <input type="text" size="6" maxlength="6" name="pekerja">
        &nbsp;&nbsp;&nbsp;&nbsp; <input type="submit" name="proses" value="Tambah"></td>
    </tr>
  </form>
  <tr class="hd"> 
    <td colspan="6" >Senarai Kakitangan Terlibat</td>
  </tr>
  <tr style="BACKGROUND-COLOR: #ffffff">
    <td colspan="6">&nbsp;</td>
  </tr>
  <tr class="hd"> 
    <td align="center">Bil</td>
    <td align="center">No Pekerja</td>
    <td align="center">Nama</td>
    <td align="center"> Anggaran OT (RM)</td>
    <td align="center">Proses</td>
  </tr>
  <%  
bil=0
jum_ot_semua=0 
'----------------------------------------filter by melebihi 1/3 gaji-----------------------------------------------------------------
q7="select pekerja_ot.rowid,paymas.nama as nama, paymas.no_pekerja,paymas.gaji_pokok as gaji_pokok,round(((paymas.gaji_pokok/3)*90/100),2) as sepertiga,"
q7 = q7 & "nvl(angg_ot_S,0) + nvl(angg_ot_M,0) as angg_ot from payroll.pekerja_ot,payroll.paymas where "
q7 = q7 & "paymas.no_pekerja=pekerja_ot.no_pekerja and to_char(tkh_ot,'mmddyyyy') like ('"& tkh_ot &"') and id_ak='"& id_ak &"' and "
q7 = q7 & "to_char(masa_tamat_angg,'HH24:MI')='"& oq6("masa_tamat") &"' and to_char(masa_mula_angg,'HH24:MI')='"& oq6("masa_mula") &"' "
q7 = q7 & "order by paymas.no_pekerja"
'if no_pekerja="9650" then response.write(q7)
set oq7 = objConn.Execute(q7)
'&" and id_ak='"& id_ak &"' 

Do While Not oq7.eof 
bil=bil + 1
jum_ot_semua = jum_ot_semua + cdbl(oq7("angg_ot"))
%>
  <form name="test<%=bil%>"  method="post" action="sp23.asp">
    <input type="hidden" value="<%=oq7("rowid")%>" name="rowid_pekerja_ot">
    <input type="hidden" value="<%=oq7("no_pekerja")%>" name="pekerja">
    <input type="hidden" name="id_ak" value="<%=id_ak%>">
    <input type="hidden" name="tkh_ot" value="<%=tkh_ot%>">
    <input type="hidden" name="rowid_jadual_ot" value="<%=rowid_jadual_ot%>">
    <% 
color="#EFF5F5"
'check sepertiga
		qs3="select nvl(sum(angg_ot_S),0) + nvl(sum(angg_ot_M),0) as total_ot_sebulan from payroll.pekerja_ot where "
		qs3 = qs3 & "no_pekerja="& oq7("no_pekerja") &" and to_char(tkh_ot,'mmyyyy') like '"& mid(tkh_ot,1,2) & mid(tkh_ot,5,4) &"'"
		'if no_pekerja="9650" then response.write(qs3)
		'response.write("<br><br>qs3---->"& qs3)
		set oqs3 = objConn.Execute(qs3)
		if not oqs3.bof and not oqs3.eof then
			'response.write("<br>total sebulan + terkini" & oqs3("total_ot_sebulan"))		
			'blh insert sbb belum melebihi 1/3 gaji
			if cdbl(oqs3("total_ot_sebulan")) > cdbl(oq7("sepertiga")) then
			color = "yellow"

				%>
<!--<script language="javascript">
	alert("Kakitangan <%'=pekerja%> sudah membuat tuntutan lebih dari 20% gaji pada bulan <%'=bulantahun%> ");
	</script>
--><%
			end if
		end if
%>
    <tr bgcolor="<%=color%>"> 
      <td><%=bil%>&nbsp;</td>
      <td><%=oq7("no_pekerja")%>&nbsp;</td>
      <td ><%=oq7("nama")%>&nbsp;</td>
      <td><div align="right"><%=formatnumber(oq7("angg_ot"),2)%></div></td>
      <td><input type="submit" value="Hapus" name="proses" onClick="return confirm('Anda Pasti Hapus Rekod?')"></td>
    </tr>
  </form>
  <%
  oq7.movenext
  loop 
'-----
oq7.close
objConn.close
'------
%>
  <tr bgcolor="#EFF5F5"> 
    <td colspan="3"><div align="right"><b>Jumlah Keseluruhan Anggaran Lebih Masa</b></div></td>
    <td><div align="right"><%=formatnumber(jum_ot_semua,2)%></div></td>
    <td>&nbsp;</td>
  </tr></tbody>
</table>
<br>
<% end sub %>
</body>
</html>