<!--#include file="connection.asp" -->
<%response.buffer=true%>
<%response.cookies("amenu") = "sp31.asp"%>
<!--'#include file="spmenu.asp"-->
<html>
<head>
<title>Sistem Pengurusan OT</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" href="css/menu1.css" REL="stylesheet">
<script type="text/javascript" src="menu1.js"></script>
<script language="javascript">
function checkMasaOT(f)
{
	if (f.mula_otH.value=="")
	{
		alert("Sila pilih jam masa masuk OT");
		f.mula_otH.focus();
		return false;
	}
	
	if (f.mula_otM.value=="")
	{
		alert("Sila pilih minit masa masuk OT");
		f.mula_otM.focus();
		return false;
	}
	if (f.tamat_otH.value=="")
	{
		alert("Sila pilih jam masa balik OT");
		f.tamat_otH.focus();
		return false;
	}
	if (f.tamat_otM.value=="")
	{
		alert("Sila pilih minit masa balik OT");
		f.tamat_otM.focus();
		return false;
	}	
	sessiMasuk = parseFloat(f.masa_mula_display.value);
	sessiBalik = parseFloat(f.masa_tamat_display.value);
	scanMasuk = parseFloat(f.mula_otH.value.concat(".",f.mula_otM.value));
	scanBalik = parseFloat(f.tamat_otH.value.concat(".",f.tamat_otM.value));
	
	if (scanMasuk > scanBalik)
	{
		alert("Pastikan masuk OT tidak melebihi dr masa balik OT");
		f.mula_otH.focus();
		return false;
	}
	
	if (scanMasuk < sessiMasuk)
	{
		alert("Pastikan masa masuk OT tidak kurang dr sessi mula OT");
		f.mula_otH.focus();
		return false;
	}
		if (scanBalik > sessiBalik)
	{
		alert("Pastikan masa balik OT tidak lebih dr sessi tamat OT");
		f.tamat_otH.focus();
		return false;
	}
}
</script>
</head>
<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%'=color4%>">
<% 
proses1=request.form("proses1")
if proses1 <> "" then 
	proses=proses1
else
	proses = request.form("proses")
end if
id_ak = request.form("id_ak")
pekerja=request.form("pekerja")
ptj=plokasi
papar
if proses <> "" then	
		if proses="Hantar" then 
			borang
		elseif proses="Simpan" then 
			gaji_pokok_pekerja=request.form("gaji_pokok")
			sepertiga=request.form("sepertiga")
			simpan
			borang
		elseif proses="Batal" then
			batal
			borang
		end if
end if

sub batal
masa_mula_asal=request.form("masa_mula")
masa_tamat_asal=request.form("masa_tamat")
'delete from proses_OT
d1="delete from payroll.proses_OT where id_ak='"& id_ak &"' and no_pekerja="& pekerja &" and "
d1 = d1 & "masa_mula_angg=to_date('"& masa_mula_asal &"','mm/dd/yyyy HH24:MI') "
d1 = d1 & "and masa_tamat_angg=to_date('"& masa_tamat_asal &"','mm/dd/yyyy HH24:MI') and kelulusan is null"
objConn.Execute(d1)

'update status_claim=null
u1="update payroll.pekerja_ot set status_claim=null,masa_masuk=null,masa_balik=null,ot_sebenar_m=null,ot_sebenar_s=null "
u1 = u1 & "where id_ak='"& id_ak &"' and no_pekerja="& pekerja &" and "
u1 = u1 & "masa_mula_angg=to_date('"& masa_mula_asal &"','mm/dd/yyyy HH24:MI') "
u1 = u1 & "and masa_tamat_angg=to_date('"& masa_tamat_asal &"','mm/dd/yyyy HH24:MI')"
objConn.Execute(u1)
end sub

sub insertTKH(id_ak,tkh,masa_mula_angg,masa_tamat_angg,sessi_bekerja,hadir1,hadir2,masa_mula,masa_tamat,kategori_masa,kategori_hari,kadar,jam,minit)
jamtotal=Cdbl(jam) + Cdbl(minit)
amaun_sepertiga=""


'check status sepertiga jika blh dibenarkan insert - to do later
'check sepertiga. if lebih x blh insert
c1="select nvl(sum(tuntutan),0) + round((("& gaji_pokok_pekerja &" * 12/2504)*"& kadar &")*"& jamtotal &",2) as total_ot_sebulan "
c1 = c1 & "from payroll.proses_OT where no_pekerja="& pekerja &" and to_char(tkh_ot,'mmyyyy') like '"& mid(tkh,1,2) & mid(tkh,7,4) &"'"
set oc1 = objConn.Execute(c1)

if not oc1.bof and not oc1.eof then
if cdbl(oc1("total_ot_sebulan")) <= cdbl(sepertiga) then statusinsert="blh"

end if


if statusinsert="blh" then

'delete utk elak duplicate rekod dlm table payroll.proses_OT
d1="delete from payroll.proses_OT where id_ak='"& id_ak &"' and no_pekerja="& pekerja &" "
d1 = d1 & "and to_char(tkh_ot,'mm/dd/yyyy') like '"& tkh &"' "
d1 = d1 & "and to_char(masa_tamat_angg,'mm/dd/yyyy HH24:MI') like '"& masa_tamat_angg &"' "
d1 = d1 & "and to_char(masa_mula_angg,'mm/dd/yyyy HH24:MI') like '"& masa_mula_angg &"' and kelulusan is null"
objConn.Execute(d1)

'-------------------------------------------------------check elaun tanggung kerja-----------------------------------
'tarikh_mula="01" & mid(tkh,1,2) & mid(tkh,7,4)
'bln=mid(tkh,1,2)
'if bln="01" or bln="03" or bln="05" or bln="07" or bln="08" or bln="10" or bln="12" then hari1="31"
'if bln="04" or bln="06" or bln="09" or bln="11" then hari1="30"
'if bln="02" then 
'thnlompat=thn Mod 4
'if thnlompat="0" then 
'hari1="29"
'else
'hari1="28"
'end if
'end if	
'tarikh_tamat=hari1 & bln & thn
'bilp=0

'qtk="select no_pekerja from elaun_mas where no_pekerja ="& pekerja &" and kod_elaun in (105,124) and "&_
'" (((tarikh_mula between to_date('"& tarikh_mula &"','ddmmyyyy') and to_date('"& tarikh_tamat &"','ddmmyyyy')) "&_
'" or (tarikh_tamat between to_date('"& tarikh_mula &"','ddmmyyyy') and to_date('"& tarikh_tamat &"','ddmmyyyy'))) "&_
'" or ((to_date('"& tarikh_mula &"','ddmmyyyy') between tarikh_mula and tarikh_tamat)"&_
'" or (to_date('"& tarikh_mula &"','ddmmyyyy') between tarikh_mula and tarikh_tamat)))"

	'set rs6=objConn.Execute(qtk)
	'if not rs6.eof then
		'a11="select count(*) bilp from tuntutan_harian where no_pekerja = "& pekerja &" and "
		'a11 = a11 & "tarikh=to_date('"&tkh&"','mm/dd/yyyy')"
		'if not q11.eof then bilp=cint(q11("bilp"))
		'if bil=1 then jamtotal=jamtotal-2.25
	'end if
	
'-------------------------tutup checking elaun tanggungkerja-----------------------------------------
	
'insert record 
i1="insert into payroll.proses_OT(id_ak,no_pekerja,bln,thn,ptj,tkh_ot,"
i1 = i1 & "masa_mula_angg,masa_tamat_angg,"
i1 = i1 & "sessi_bekerja,hadir1,hadir2,masa_mula,masa_tamat,kategori_masa,"
i1 = i1 & "kategori_hari,kadar,jam,minit,gaji_pokok,tuntutan) values "
i1 = i1 & "('"& id_ak &"',"& pekerja &","& mid(tkh,1,2) &","& mid(tkh,7,4) &","& ptj &",to_date('"& tkh &"','mm/dd/yyyy'),"
i1 = i1 & "to_date('"& masa_mula_angg &"','mm/dd/yyyy HH24:MI'),to_date('"& masa_tamat_angg &"','mm/dd/yyyy HH24:MI'),"
i1 = i1 & "'"& sessi_bekerja &"','"& hadir1 &"','"& hadir2 &"',"
i1 = i1 & "to_date('"& masa_mula &"','mm/dd/yyyy HH24:MI'),to_date('"& masa_tamat &"','mm/dd/yyyy HH24:MI'),"
i1 = i1 & "'"& kategori_masa &"','"& kategori_hari &"',"& kadar &","& jam &",round("& minit &",2),"& gaji_pokok_pekerja &","
i1 = i1 & "round((("& gaji_pokok_pekerja &" * 12/2504)*"& kadar &")*"& jamtotal &",2))"
objConn.Execute(i1)

o="update payroll.pekerja_ot set masa_masuk=to_date('"& masa_mula &"','mm/dd/yyyy HH24:MI'),"
o = o & "masa_balik=to_date('"& masa_tamat &"','mm/dd/yyyy HH24:MI'),"
if kategori_masa="S" then
o = o & "ot_sebenar_s=round((("& gaji_pokok_pekerja &" * 12/2504)*"& kadar &")*"& jamtotal &",2) "
else
o = o & "ot_sebenar_m=round((("& gaji_pokok_pekerja &" * 12/2504)*"& kadar &")*"& jamtotal &",2) "
end if
o = o & "where id_ak='"& id_ak &"' and no_pekerja="& pekerja & " and "
o = o & "masa_mula_angg=to_date('"& masa_mula_angg &"','mm/dd/yyyy HH24:MI') "
o = o & "and masa_tamat_angg=to_date('"& masa_tamat_angg &"','mm/dd/yyyy HH24:MI')"
objConn.Execute(o)
else
%>
	<script language="javascript">
	alert("Rekod pekerja <%=pekerja%> sudah melebihi 1/3 gaji");
	</script>
<%
end if
end sub

sub simpan ()
check1 ="select distinct id_ak from payroll.pekerja_ot where no_pekerja="& pekerja &" and sah_individu is not null"
set qcheck1 = objConn.Execute(check1)

if not qcheck1.bof and not qcheck1.eof then
%>
	<script language="javascript">
	alert("Rekod tuntutan OT pekerja <%=pekerja%> sudah disahkan penyelia\n.Proses tidak dpt diteruskan");
	</script>
<%
else
rowidot=request.form("rowidot")
kategori_hari=request.form("kategori_hari")
tkh_ot=request.form("tkh_ot") 'mm/dd/yyyy
masa_mula_asal=request.form("masa_mula")
masa_tamat_asal=request.form("masa_tamat")
masa_mula_display=request.form("masa_mula_display")
masa_tamat_display=request.form("masa_tamat_display")
sessi_bekerja=request.form("sessi_bekerja")
hadir1=request.form("hadir1")
hadir2=request.form("hadir2")

masa_mulaKira = request.form("mula_otH")&"."& request.form("mula_otM")
masa_tamatKira = request.form("tamat_otH")&"."& request.form("tamat_otM")
masa_mula = tkh_ot &" "& replace(masa_mulaKira,".",":")
masa_tamat = tkh_ot &" "& replace(masa_tamatKira,".",":")


'insert into proses_tuntutan
if Cdbl(masa_mulaKira)>=6 and Cdbl(masa_tamatKira) <=22 then
	'cth 6.00 -> 22.00
	'status1="S"
	'status2="S"
	
	'cari beza masa
	jum_minit = datediff("n", Cdate(masa_mula), Cdate(masa_tamat))
	jam = Fix(jum_minit / 60)
	minit = (jum_minit mod 60)/60
	
	if kategori_hari="P" then	
		kadar=1.125
	elseif kategori_hari="SA" then
		kadar=1.25
	elseif kategori_hari="CU" then
		kadar=1.75
	end if
	'call sub utk insert record
	call insertTkh(id_ak,tkh_ot,masa_mula_asal,masa_tamat_asal,sessi_bekerja,hadir1,hadir2,replace(masa_mula,".",":"),replace(masa_tamat,".",":"),"S",kategori_hari,kadar,jam,minit)
	
elseif Cdbl(masa_mulaKira)<6 and Cdbl(masa_tamatKira) > 22 then
	'cth=5.00 -> 23.00
	'status1="M"
	'status2="S"
	'status3="M"
	
	'cari beza masa 
	jum_minit_status1 = datediff("n", Cdate(masa_mula), Cdate(tkh_ot &" 06:00"))
	jum_minit_status3 = datediff("n", Cdate(tkh_ot &" 22:00"), Cdate(masa_tamat))
	jum_minit = Cint(jum_minit_status1) + Cint(jum_minit_status3)
	'response.write("<br>--->jum minit malam"& jum_minit)
	jam_M = Fix(jum_minit / 60)
	minit_M = (jum_minit mod 60)/60
	jum_minit = datediff("n", Cdate(tkh_ot &" 06:00"), Cdate(tkh_ot &" 22:00"))
	'response.write("<br>--->jum minit siang"& jum_minit)
	jam_S = Fix(jum_minit / 60)
	minit_S = (jum_minit mod 60)/60
	
	if kategori_hari="P" then
		kadar_M=1.25
		kadar_S=1.125
	elseif kategori_hari="SA" then
		kadar_S=1.25
		kadar_M=1.50
	elseif kategori_hari="CU" then
		kadar_S=1.75
		kadar_M=2.00
	end if

	'call sub utk insert record 'ada 2 kadar, kadar Siang dan kadar Malam
	call insertTkh(id_ak,tkh_ot,masa_mula_asal,masa_tamat_asal,sessi_bekerja,hadir1,hadir2,replace(masa_mula,".",":"),replace(masa_tamat,".",":"),"S",kategori_hari,kadar_S,jam_S,minit_S)
	call insertTkh(id_ak,tkh_ot,masa_mula_asal,masa_tamat_asal,sessi_bekerja,hadir1,hadir2,replace(masa_mula,".",":"),replace(masa_tamat,".",":"),"M",kategori_hari,kadar_M,jam_M,minit_M)	

	
elseif Cdbl(masa_mulaKira)>22 and Cdbl(masa_tamatKira) > 22 then
	'cth 22.30 -> 24.00
	'status1="M"
	'status2="M"
	
	'cari beza masa
	'response.write("<br>masa_mula--->"& masa_mula&"---->"& len(masa_mula))
	'response.write("<br>masa_tamat--->"& masa_tamat&"---->"& len(masa_tamat))
	
	jum_minit = datediff("n", Cdate(masa_mula), Cdate(masa_tamat))
	
	jam_M = Fix(jum_minit / 60)
	minit_M = (jum_minit mod 60)/60	
	
	if kategori_hari="P" then
		kadar_M=1.25
	elseif kategori_hari="SA" then
		kadar_M=1.50
	elseif kategori_hari="CU" then
		kadar_M=2.00
	end if
	
	'call sub utk insert record kadar Malam
	call insertTkh(id_ak,tkh_ot,masa_mula_asal,masa_tamat_asal,sessi_bekerja,hadir1,hadir2,replace(masa_mula,".",":"),replace(masa_tamat,".",":"),"M",kategori_hari,kadar_M,jam_M,minit_M)	
	
elseif Cdbl(masa_mulaKira)<6 and Cdbl(masa_tamatKira) < 6 then
	'cth 01.00 -> 4.00
	'status1="M"
	'status2="M"
	
	'cari beza masa
	jum_minit = datediff("n", Cdate(masa_mula), Cdate(masa_tamat))
	jam_M = Fix(jum_minit / 60)
	minit_M = (jum_minit mod 60)/60	
	
	if kategori_hari="P" then
		kadar_M=1.25
	elseif kategori_hari="SA" then
		kadar_M=1.50
	elseif kategori_hari="CU" then
		kadar_M=2.00
	end if
	
	'call sub utk insert record kadar Malam
	call insertTkh(id_ak,tkh_ot,masa_mula_asal,masa_tamat_asal,sessi_bekerja,hadir1,hadir2,replace(masa_mula,".",":"),replace(masa_tamat,".",":"),"M",kategori_hari,kadar_M,jam_M,minit_M)	
	
elseif Cdbl(masa_mulaKira)>=6 and Cdbl(masa_tamatKira) > 22 then
	'cth 017.00 -> 23.00
	'status1="S"
	'status2="M"
	
	'cari beza masa
	jum_minit = datediff("n", Cdate(masa_mula), Cdate(tkh_ot &" 22:00"))
	jam_S = Fix(jum_minit / 60)
	minit_S = (jum_minit mod 60)/60
	
	jum_minit = datediff("n", Cdate(tkh_ot &" 22:00"), Cdate(masa_tamat)) 'Asma
	jam_M = Fix(jum_minit / 60)
	minit_M = (jum_minit mod 60)/60	
	
	if kategori_hari="P" then
		kadar_M=1.25
		kadar_S=1.125
	elseif kategori_hari="SA" then
		kadar_S=1.25
		kadar_M=1.50
	elseif kategori_hari="CU" then
		kadar_S=1.75
		kadar_M=2.00
	end if
	
	'call sub utk insert record 'ada 2 kadar, kadar Siang dan kadar Malam
	call insertTkh(id_ak,tkh_ot,masa_mula_asal,masa_tamat_asal,sessi_bekerja,hadir1,hadir2,replace(masa_mula,".",":"),replace(masa_tamat,".",":"),"S",kategori_hari,kadar_S,jam_S,minit_S)
	call insertTkh(id_ak,tkh_ot,masa_mula_asal,masa_tamat_asal,sessi_bekerja,hadir1,hadir2,replace(masa_mula,".",":"),replace(masa_tamat,".",":"),"M",kategori_hari,kadar_M,jam_M,minit_M)	
	
elseif Cdbl(masa_mulaKira)<6 and Cdbl(masa_tamatKira) <= 22 then

	'cth 05.00 -> 18.00
	'status1="M"
	'status2="S"
	
	'cari beza masa
	jum_minit = datediff("n", Cdate(masa_mula), Cdate(tkh &" 06:00"))
	jam_M = Fix(jum_minit / 60)
	minit_M = (jum_minit mod 60)/60	
	
	jum_minit = datediff("n", Cdate(tkh_ot &" 06:00"), Cdate(masa_tamat))
	jam_S = Fix(jum_minit / 60)
	minit_S = (jum_minit mod 60)/60
	
	if kategori_hari="P" then
		kadar_M=1.25
		kadar_S=1.125
	elseif kategori_hari="SA" then
		kadar_S=1.25
		kadar_M=1.50
	elseif kategori_hari="CU" then
		kadar_S=1.75
		kadar_M=2.00
	end if	
	
	'call sub utk insert record 'ada 2 kadar, kadar Siang dan kadar Malam
	call insertTkh(id_ak,tkh_ot,masa_mula_asal,masa_tamat_asal,sessi_bekerja,hadir1,hadir2,replace(masa_mula,".",":"),replace(masa_tamat,".",":"),"S",kategori_hari,kadar_S,jam_S,minit_S)
	call insertTkh(id_ak,tkh_ot,masa_mula_asal,masa_tamat_asal,sessi_bekerja,hadir1,hadir2,replace(masa_mula,".",":"),replace(masa_tamat,".",":"),"M",kategori_hari,kadar_M,jam_M,minit_M)
else
	response.write("Proses tidak lengkap bagi masa "& masa_mula &" hingga "& masa_tamat &"<br>Sila hubungi system administrator<br><br>")

end if


'update status ='Y' in pekerja_ot
u1="update payroll.pekerja_ot set status_claim='Y' where id_ak='"& id_ak &"' and no_pekerja="& pekerja &" and "
u1 = u1 & "masa_mula_angg=to_date('"& masa_mula_asal &"','mm/dd/yyyy HH24:MI') "
u1 = u1 & "and masa_tamat_angg=to_date('"& masa_tamat_asal &"','mm/dd/yyyy HH24:MI')"
objConn.Execute(u1)
end if 'eof checking sah_individu is not null
end sub

sub papar ()
q1="select id_ak from payroll.arahan_kerja where ptj="& ptj &" and tkh_lulus is null and tkh_sah is not null order by tahun desc,bulan desc,id_ak"
set oq1 = objConn.Execute(q1)
%>
<br>
<form name="test" method="post" action="sp31.asp">
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
<tr class="hd">
  <td colspan="2">Semak/Input Tuntutan OT Seorang Pekerja Oleh Penyelia</td></tr>
<tr>
  <td width="15%"><b> No Rujukan Arahan Kerja:</b> </td>
  <td><select name="id_ak">
  <option value="">Sila Pilih</option>
  <%do while not oq1.eof %>
    <option <%if id_ak=oq1("id_ak") then %> selected <%end if%>value="<%=oq1("id_ak")%>"><%=oq1("id_ak")%></option>
    <%oq1.movenext
	loop%>
  </select></td></tr>
<tr>
  <td><b>No Pekerja</b></td>
  <td><input type="text" size="6" maxlength="6" name="pekerja"></td>
</tr>
<tr align="center"><td colspan="2"><input type="submit" name="proses" value="Hantar"></td></tr>
</table>
</form>
<%
end sub %>
<%

sub borang()
q0="select * from payroll.pekerja_ot where id_ak='"& id_ak &"' and no_pekerja="& pekerja &""
set oq0 = objConn.Execute(q0)
if oq0.bof and oq0.eof then
%>
<br>
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
  <TBODY>
    <tr> <td><div align="center">Tiada rekod bagi pekerja <%=pekerja%></div></td></tr></table>
<%
else
q1="select penyelia,pemohon,pengesahan,unit,ptj,keterangan,nota,bulan,tahun from payroll.arahan_kerja where id_ak ='"& id_ak &"'"
set oq1 = objConn.Execute(q1)

tkh_input=oq1("bulan") &"/01/"& oq1("tahun")
TodayDate = DatePart("m",Date) &"/01/"& DatePart("yyyy",Date)

if Cdate(TodayDate) < Cdate(tkh_input) then
		q2="select upper(nama) as nama, amaun as gaji_pokok, amaun/10 as sepertiga from payroll.paymas,payroll.paymas_txn where "
		q2 = q2 & "paymas_txn.bulan="& oq1("bulan") &" and no_pekerja="& pekerja &" and paymas_txn.tahun="& oq1("tahun") &" and "
		q2 = q2 & "paymas_txn.no_pekerja=paymas.no_pekerja"
		set oq2 = objConn.Execute(q2)
		If not oq2.bof and not oq2.eof then
			nama = oq2("nama")
			gaji_pokok = oq2("gaji_pokok")
			sepertiga = oq2("sepertiga")
		end if
else
	q2="select upper(nama) as nama, gaji_pokok, round(gaji_pokok/10,2) as sepertiga from payroll.paymas where no_pekerja="& pekerja &""
		set oq2 = objConn.Execute(q2)
		If not oq2.bof and not oq2.eof then
			nama = oq2("nama")
			gaji_pokok = oq2("gaji_pokok")
			sepertiga = oq2("sepertiga")
		end if
end if
q2="select upper(nama) as nama from payroll.paymas where no_pekerja="& oq1("penyelia") &""
	set oq2 = objConn.Execute(q2)
	If not oq2.bof and not oq2.eof then
		penyelia = oq2("nama")
	end if
q2="select upper(nama) as nama from payroll.paymas where no_pekerja="& oq1("pemohon") &""
	set oq2 = objConn.Execute(q2)
	If not oq2.bof and not oq2.eof then
		pemohon = oq2("nama")
	end if
%>
<br>
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
  <TBODY>
    <tr class="hd"> 
      <td colspan="11" >No Pekerja : <%=pekerja%><br>
      Nama Pekerja : <%=nama%><br>
      <%'Gaji 1/3 : <%=sepertiga<br>%>
      No Rujukan Arahan Kerja : <%=id_ak%><br>
      Bln/Thn : <%=oq1("bulan")%> /<%=oq1("tahun")%><br>
      Penyelia : <%=penyelia%><br>
      Pemohon: <%=pemohon%><br>
      Keterangan : <%=oq1("keterangan")%><br>
      Nota Tambahan : <%=oq1("nota")%><br>
      </td>
    </tr> 
  <tr style="BACKGROUND-COLOR: #ffffff"><td colspan="11">&nbsp;</td></tr>
  <tr class="hd"> 
    <td rowspan="2" align="center">Bil</td>
    <td rowspan="2" align="center">Tarikh OT</td>
    <td rowspan="2" align="center">Sessi  Bekerja </td>
      <td colspan="2" align="center">Rekod Kedatangan</td>
      <td colspan="2" align="center">Waktu Arahan Kerja</td>
      <td colspan="2" align="center">Tuntutan Lebih Masa</td>
    <td rowspan="2" align="center">Proses</td>
  </tr>
  <tr class="hd">
    <td align="center">Sessi1</td>
    <td align="center">Sessi2</td>
    <td align="center">&nbsp;</td>
    <td align="center">&nbsp;</td>
    <td align="center">Masuk</td>
    <td align="center">Balik</td>
  </tr>
  <% 
'-----------------set array-----
dim datakehadiran(30,5)
'0=tkh ot
'1=sessi bekerja
'2=masa masuk
'3=masa keluar

'get jenis sessi (shift/normal)
qa="select count(*) bilangan from kehadiran.jadual where no_pekerja="& pekerja &" and tahun="& oq1("tahun") &" "
set oqa = objekConn.Execute(qa)
if oqa("bilangan") = 0 then
	status_sessi="tiada"
	sessi_bulanan="0"
elseif oqa("bilangan") = 1 then
	status_sessi="normal"
	qb="select nvl(B"& replace(oq1("bulan"),"0","") &",0) as sessibln,waktu_bekerja.masa_mula as smula,waktu_bekerja.masa_tamat as stamat "
	qb = qb & "from kehadiran.jadual,kehadiran.waktu_bekerja where "
	qb = qb & "no_pekerja="& pekerja &" and tahun="& oq1("tahun") &" and waktu_bekerja.kod=nvl(jadual.B"& replace(oq1("bulan"),"0","") &",0)"
	set oqb=objekConn.Execute(qb)
	if not oqb.bof and not oqb.eof then
		sessi_bulanan=oqb("smula") &" - "& oqb("stamat")
	else
		sessi_bulanan="0"
	end if
else	
	status_sessi="shift"
end if

q3="select to_char(tkh_ot,'mmddyyyy') as tkh_ot,to_char(masa_mula,'HH24:MI') as masa_mula,"
q3 = q3 & "to_char(masa_tamat,'HH24:MI') as masa_tamat,masa_mula as susun "
q3 = q3 & "from payroll.jadual_ot where id_ak ='"& id_ak &"' and "
q3 = q3 & "masa_mula in (select masa_mula_angg from payroll.pekerja_ot where id_ak='"& id_ak &"' and no_pekerja="& pekerja &") "
q3 = q3 & "order by tkh_ot asc,susun"
set oq3 = objConn.Execute(q3)  
bil_rekod=0
do while not oq3.eof
	datakehadiran(bil_rekod,0)=oq3("tkh_ot")
	'set session
	if status_sessi="shift" then
		'get session shift
		qb="select SESI_BEKERJA as sessibln,waktu_bekerja.masa_mula as smula,waktu_bekerja.masa_tamat as stamat "
		qb = qb & "from kehadiran.jadual,kehadiran.waktu_bekerja where "
		qb = qb & "no_pekerja="& pekerja &" and tahun="& oq1("tahun") &" and "
		qb = qb & "to_date('"& oq3("tkh_ot") &"','mmddyyyy') between tarikh_mula and tarikh_tamat "
		qb = qb & "and waktu_bekerja.kod=jadual.SESI_BEKERJA"
		set oqb=objekConn.Execute(qb)
		if not oqb.bof and not oqb.eof then
			sessi_bulanan=oqb("smula") &" - "& oqb("stamat")
		else
			sessi_bulanan="0"
		end if
	end if
	datakehadiran(bil_rekod,1)=sessi_bulanan
	
	'get kehadiran
	q4="select nvl(time1,'0') time1,nvl(time2,'0') time2,nvl(time3,'0') time3,nvl(time4,'0') time4,komen from kehadiran.kehadiran where "
	q4 = q4 & "no_pekerja="& pekerja &" and tarikh like to_date('"& oq3("tkh_ot") &"','mmddyyyy')"
	set oq4 = objekConn.Execute(q4)
	if not oq4.bof and not oq4.eof then
		datakehadiran(bil_rekod,2)=oq4("time1")
		datakehadiran(bil_rekod,3)=oq4("time2")
		datakehadiran(bil_rekod,4)=oq4("time3")
		datakehadiran(bil_rekod,5)=oq4("time4")		
	else
		datakehadiran(bil_rekod,2)="0"
		datakehadiran(bil_rekod,3)="0"
		datakehadiran(bil_rekod,4)="0"
		datakehadiran(bil_rekod,5)="0"	
	end if
	
bil_rekod = bil_rekod + 1
oq3.movenext
loop
'-----------------------------------------------------------

bil=0
bild=0
q3="select to_char(tkh_ot,'dd/mm/yyyy') as tkh_ot_display,to_char(tkh_ot,'mm/dd/yyyy') as tkh_ot_value,"
q3 = q3 & "to_char(masa_mula,'mm/dd/yyyy HH24.MI') as masa_mula,to_char(masa_tamat,'mm/dd/yyyy HH24.MI') as masa_tamat,"
q3 = q3 & "to_char(masa_mula,'HH24:MI') as masa_mula_display,masa_mula as susun,rowid,kategori_hari,"
q3 = q3 &" decode(to_char(masa_tamat,'HH24:MI'),'00:00','24:00',to_char(masa_tamat,'HH24:MI')) as masa_tamat_display "
q3 = q3 & "from payroll.jadual_ot where id_ak ='"& id_ak &"' and "
q3 = q3 & "masa_mula in (select masa_mula_angg from payroll.pekerja_ot where id_ak='"& id_ak &"' and no_pekerja="& pekerja &") order by tkh_ot asc,susun"
set oq3 = objConn.Execute(q3)
bil=0

do while not oq3.eof
bild=bild+1
mula_ot=-1
tamat_ot=-1
masuk1=-1
masuk2=-1
balik1=-1
balik2=-1
jam_masuk_ot=-1
minit_masuk_ot=-1

jam_balik_ot=-1
minit_balik_ot=-1

mula_ot = replace(oq3("masa_mula_display"),":",".") 'masa ot
tamat_ot = replace(oq3("masa_tamat_display"),":",".")
if datakehadiran(bil,2) <> "" then masuk1=mid(datakehadiran(bil,2),1,2) &"."& mid(datakehadiran(bil,2),3,2) 'masa kehadiran
if datakehadiran(bil,3) <> "" then balik1=mid(datakehadiran(bil,3),1,2) &"."& mid(datakehadiran(bil,3),3,2)

'check satu sessi atau 2 sessi
if datakehadiran(bil,2) <> "0" and datakehadiran(bil,4) <> "0" then '2 sessi 

	mula_sessi=-1
	tamat_sessi=-1
	
	mula_sessi=replace(mid(datakehadiran(bil,1),1,5),":",".") 'masa sessi
	'tamat_sessi=replace(mid(datakehadiran(bil,1),9,5),":",".")
	
	'check sessi bekerja
	if cdbl(mula_sessi) < cdbl(mula_ot) then 'sessi bekerja pagi.  sessi ot ambil dr sessi 2
	
		masuk2=mid(datakehadiran(bil,4),1,2) &"."& mid(datakehadiran(bil,4),3,2) 'masa kehadiran sessi1
		balik2=mid(datakehadiran(bil,5),1,2) &"."& mid(datakehadiran(bil,5),3,2)		
		
		if Cdbl(masuk2) >= Cdbl(mula_ot) then
			jam_masuk_ot=mid(masuk2,1,InStr(masuk2,".")-1)
			minit_masuk_ot=mid(masuk2,InStr(masuk2,".")+1,len(masuk1))
		else
			jam_masuk_ot=mid(mula_ot,1,InStr(mula_ot,".")-1)
			minit_masuk_ot=mid(mula_ot,InStr(mula_ot,".")+1,len(mula_ot))
		end if
		'set masa balik ot
		if cdbl(balik2) <= cdbl(tamat_ot) then
			jam_balik_ot = mid(balik2,1,InStr(balik2,".")-1)
			minit_balik_ot = mid(balik2,InStr(balik2,".")+1,len(balik1))
		else
			jam_balik_ot = mid(tamat_ot,1,InStr(tamat_ot,".")-1)
			minit_balik_ot = mid(tamat_ot,InStr(tamat_ot,".")+1,len(tamat_ot))
		end if
		
	else 'sessi bekerja ptg.  sessi ot ambil dr sessi 1		
		if Cdbl(masuk1) >= Cdbl(mula_ot) then
			jam_masuk_ot=mid(masuk1,1,InStr(masuk1,".")-1)
			minit_masuk_ot=mid(masuk1,InStr(masuk1,".")+1,len(masuk1))
		else
			jam_masuk_ot=mid(mula_ot,1,InStr(mula_ot,".")-1)
			minit_masuk_ot=mid(mula_ot,InStr(mula_ot,".")+1,len(mula_ot))
		end if
		'set masa balik ot
		if cdbl(balik1) <= cdbl(tamat_ot) then
			jam_balik_ot = mid(balik1,1,InStr(balik1,".")-1)
			minit_balik_ot = mid(balik1,InStr(balik1,".")+1,len(balik1))
		else
			jam_balik_ot = mid(tamat_ot,1,InStr(tamat_ot,".")-1)
			minit_balik_ot = mid(tamat_ot,InStr(tamat_ot,".")+1,len(tamat_ot))
		end if
	end if
	
	
elseif datakehadiran(bil,2) <> "0" and datakehadiran(bil,4) = "0" then   '1 sessi
	'set masa mula ot
	if Cdbl(masuk1) >= Cdbl(mula_ot) then
		jam_masuk_ot=mid(masuk1,1,InStr(masuk1,".")-1)
		minit_masuk_ot=mid(masuk1,InStr(masuk1,".")+1,len(masuk1))
	else
		jam_masuk_ot=mid(mula_ot,1,InStr(mula_ot,".")-1)
		minit_masuk_ot=mid(mula_ot,InStr(mula_ot,".")+1,len(mula_ot))
	end if
	'set masa balik ot
	if cdbl(balik1) <= cdbl(tamat_ot) then
		jam_balik_ot = mid(balik1,1,InStr(balik1,".")-1)
		minit_balik_ot = mid(balik1,InStr(balik1,".")+1,len(balik1))
	else
		jam_balik_ot = mid(tamat_ot,1,InStr(tamat_ot,".")-1)
		minit_balik_ot = mid(tamat_ot,InStr(tamat_ot,".")+1,len(tamat_ot))
	end if
	
	'response.write ("jam_masuk_ot--->"& jam_masuk_ot &"minit_masuk_ot--->"& minit_masuk_ot &"jam_balik_ot--->"&jam_balik_ot&"minit_balik_ot-->"& minit_balik_ot&"<br>")
	
end if 
'response.write("minit_masuk_ot-->"& minit_masuk_ot &"minit_balik_ot--->"& minit_balik_ot &"<br>")
	if Cdbl(minit_masuk_ot)=-1 then 
		minit_masuk_ot=-1
		'response.write("<br>-1")
	elseif Cdbl(minit_masuk_ot)<30 then 
		minit_masuk_ot=0
		'response.write("<br>0")
	elseif Cdbl(minit_masuk_ot)<=59 then 
		minit_masuk_ot=30
		'response.write("<br>30")
	end if
	
	if Cdbl(minit_balik_ot)=-1 then 
		minit_balik_ot=-1
		'response.write("<br>-1")
	elseif Cdbl(minit_balik_ot)<30 then 
		minit_balik_ot=0
		'response.write("<br>0")
	elseif Cdbl(minit_balik_ot)<=59 then 
		minit_balik_ot=30	
		'response.write("<br>30")	
	end if
	'response.write("2.    minit_masuk_ot-->"& minit_masuk_ot &"minit_balik_ot--->"& minit_balik_ot &"<br>")
	
'get status
statusbutton=""
statussessi=""
g1="select status_claim,to_char(masa_masuk,'HH24:MI') as masa_masuk,to_char(masa_balik,'HH24:MI') as masa_balik "
g1 = g1 & "from payroll.pekerja_ot where id_ak='"& id_ak &"' and no_pekerja="& pekerja &" and "
g1 = g1 & "masa_mula_angg=to_date('"& oq3("masa_mula") &"','mm/dd/yyyy HH24:MI') "
g1 = g1 & "and masa_tamat_angg=to_date('"& oq3("masa_tamat") &"','mm/dd/yyyy HH24:MI')"
'response.write(g1)
set og1 = objConn.Execute(g1)
if not og1.bof and not og1.eof then
	if og1("status_claim")="Y" then statusbutton="Y"
	masa_mula_sebenar=og1("masa_masuk")
	masa_tamat_sebenar=og1("masa_balik")
end if

if datakehadiran(bil,1) <> "0" and oq3("kategori_hari") = "P" then
	dd1="select 'x' from dual where (to_date('"& oq3("masa_mula") &"','mm/dd/yyyy HH24:MI') between "
	dd1 = dd1 & "to_date('"& oq3("tkh_ot_value") & mid(datakehadiran(bil,1),1,5) &"','mm/dd/yyyy HH24:MI') and "
	dd1 = dd1 & "to_date('"& oq3("tkh_ot_value") & mid(datakehadiran(bil,1),9,5) &"','mm/dd/yyyy HH24:MI')) or "
	dd1 = dd1 & "(to_date('"& oq3("masa_tamat") &"','mm/dd/yyyy HH24:MI') between "
	dd1 = dd1 & "to_date('"& oq3("tkh_ot_value") & mid(datakehadiran(bil,1),1,5) &"','mm/dd/yyyy HH24:MI') and "
	dd1 = dd1 & "to_date('"& oq3("tkh_ot_value") & mid(datakehadiran(bil,1),9,5) &"','mm/dd/yyyy HH24:MI')) "
	'response.write(dd1)
	set qdd1 = objConn.Execute(dd1)
	if not qdd1.bof and not qdd1.eof then
		statussessi="Arahan Kerja berada dalam sessi bekerja"
	end if
end if

'statushari = P,SA,CU

%>
<form name="tuntut<%=bil%>" method="post" action="sp31.asp" onSubmit="return checkMasaOT(this)">
<input type="hidden" name="pekerja" value="<%=pekerja%>">
<input type="hidden" name="id_ak" value="<%=id_ak%>">
<input type="hidden" name="kategori_hari" value="<%=oq3("kategori_hari")%>">
<input type="hidden" name="rowidot" value="<%=oq3("rowid")%>">
<input type="hidden" name="sepertiga" value="<%=sepertiga%>">
<input type="hidden" name="tkh_ot" value="<%=oq3("tkh_ot_value")%>">
<input type="hidden" name="masa_mula" value="<%=oq3("masa_mula")%>">
<input type="hidden" name="masa_tamat" value="<%=oq3("masa_tamat")%>">
<input type="hidden" name="masa_mula_display" value="<%=replace(oq3("masa_mula_display"),":",".")%>">
<input type="hidden" name="masa_tamat_display" value="<%=replace(oq3("masa_tamat_display"),":",".")%>">
<input type="hidden" name="sessi_bekerja" value="<%=datakehadiran(bil,1)%>">
<input type="hidden" name="hadir1" value="<%=datakehadiran(bil,2)%>&nbsp;-&nbsp;<%=datakehadiran(bil,3)%>">
<input type="hidden" name="hadir2" value="<%=datakehadiran(bil,4)%>&nbsp;-&nbsp;<%=datakehadiran(bil,5)%>">
<input type="hidden" name="gaji_pokok" value="<%=gaji_pokok%>">
 <%if statusbutton="Y" then%>
   <tr>
    <td><%=bild%></td>
    <td><div align="center"><%=oq3("tkh_ot_display")%></div></td>
    <td><div align="center"><%if oq3("kategori_hari")="P" then%><%=datakehadiran(bil,1)%><%else%>Hari Minggu/Cuti Umum<%end if%></div></td>
    <td align="center" ><div align="center"><%=datakehadiran(bil,2)%>&nbsp;-&nbsp;<%=datakehadiran(bil,3)%></div></td>
    <td align="center" ><div align="center"><%=datakehadiran(bil,4)%>&nbsp;-&nbsp;<%=datakehadiran(bil,5)%></div></td>
    <td align="center"><%=oq3("masa_mula_display")%></td>
    <td align="center"><%=oq3("masa_tamat_display")%></td>
    <td><%=masa_mula_sebenar%>
	<input type="hidden" name="mula_otH" value="<%=mid(masa_mula_sebenar,1,2)%>"><input type="hidden" name="mula_otM" value="<%=mid(masa_mula_sebenar,4,2)%>"></td>
    <td><%=masa_tamat_sebenar%>
	<input type="hidden" name="tamat_otH" value="<%=mid(masa_tamat_sebenar,1,2)%>"><input type="hidden" name="tamat_otM" value="<%=mid(masa_tamat_sebenar,4,2)%>"></td>
    <td><input type="submit" value="Batal" name="proses1"></td>
  </tr>
 <%else%>
  <tr>
    <td><%=bild%></td>
    <td><div align="center"><%=oq3("tkh_ot_display")%></div></td>
    <td><div align="center"><%if oq3("kategori_hari")="P" then%><%=datakehadiran(bil,1)%><%else%>Hari Minggu/Cuti Umum<%end if%></div></td>
    <td align="center" ><div align="center"><%=datakehadiran(bil,2)%>&nbsp;-&nbsp;<%=datakehadiran(bil,3)%></div></td>
    <td align="center" ><div align="center"><%=datakehadiran(bil,4)%>&nbsp;-&nbsp;<%=datakehadiran(bil,5)%></div></td>
    <td align="center"><%=oq3("masa_mula_display")%></td>
    <td align="center"><%=oq3("masa_tamat_display")%></td>
    <td>
      <select size="1" name="mula_otH">
        <%For i = 0 To 23 
	if len(i)=1 then 
		jam="0"& Cstr(i)
	else
		jam=Cstr(i)
	end if
	%>
        <option <%if cdbl(jam_masuk_ot)=cdbl(jam) then%>selected <%end if%>value="<%=jam%>"><%=jam%></option>
        <%Next%>
        </select> 
      &nbsp;:&nbsp;
      <select name="mula_otM">
        <option <%if cdbl(minit_masuk_ot)=0 then%>selected <%end if%>value="00">00</option>
  		<option <%if cdbl(minit_masuk_ot)=30 then%>selected <%end if%>value="30">30</option>
      </select></td>
    <td>
      <select name="tamat_otH">
        <%For i = 0 To 23 
	if len(i)=1 then 
		jam="0"& Cstr(i)
	else
		jam=Cstr(i)
	end if
	%>
        <option <%if cdbl(jam_balik_ot)=cdbl(jam) then%> selected <%end if%>value="<%=jam%>"><%=jam%></option>
        <%Next%>
      </select>
&nbsp;:&nbsp;
<select name="tamat_otM">
        <option <%if Cdbl(minit_balik_ot)=0 then%>selected <%end if%>value="00">00</option>
  		<option <%if Cdbl(minit_balik_ot)=30 then%>selected <%end if%>value="30">30</option>
</select></td>
    <td><%if statussessi="" then%><input type="submit" value="Simpan" name="proses"><%else%><%=statussessi%><%end if%></td>
  </tr>
  <%end if%>
  </form>
  <%
  bil=bil + 1
  oq3.movenext
  loop %>
    <tr>
    <td colspan="7">&nbsp;</td>
    <td colspan="2">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  </tbody> 
</table>
<%end if%>
<br>
<% 
end sub %>
</body>
</html>