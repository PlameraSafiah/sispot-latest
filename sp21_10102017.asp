<!--#include file="connection.asp" -->
<!--'#include file="spmenu.asp"-->
<%response.buffer=true%>
<%response.cookies("amenu") = "sp21.asp"%>
<html>
<head>
<title>Sistem Pengurusan OT</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" href="css/menu1.css" REL="stylesheet">
<script type="text/javascript" src="menu1.js"></script>
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
function checkForm1(f){
	if (f.keterangan.value=="")
	{
		alert("Sila masukkan keterangan arahan kerja");
		f.keterangan.focus();
		return false
	}
}

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

function check(f){
	if (f.tkh_bayar_gaji.value=="")
	{
		alert("Sila masukkan tarikh pembayaran gaji");
		f.tkh_bayar_gaji.focus();
		return false
	}
}

function checkTambah(f){
	if (f.id.value=="")
	{
		alert("Sila simpan maklumat arahan kerja sebelum masukkan jadual bertugas");
		return false;
	}
	if (f.tkh.value=="")
	{
		alert("Sila pilih tarikh OT");
		f.tkh.focus();
		return false
	}
		if (f.masa_mulaH.value=="")
	{
		alert("Sila pilih jam masa mula OT");
		f.masa_mulaH.focus();
		return false
	}
		if (f.masa_mulaM.value=="")
	{
		alert("Sila pilih minit masa mula OT");
		f.masa_mulaM.focus();
		return false
	}
	if (f.masa_tamatH.value=="")
	{
		alert("Sila pilih jam masa tamat OT");
		f.masa_tamatH.focus();
		return false
	}
		if (f.masa_tamatM.value=="")
	{
		alert("Sila pilih minit masa tamat OT");
		f.masa_tamatM.focus();
		return false
	}
	
	masamula = parseFloat(f.masa_mulaH.value.concat(".",f.masa_mulaM.value));
	masatamat = parseFloat(f.masa_tamatH.value.concat(".",f.masa_tamatM.value));
	if (f.masa_tamatH.value=="00")
	{
		masatamat = parseFloat(concat("24.",f.masa_tamatM.value));
	}
	
	if (masamula > masatamat)
	{
		alert("Pastikan mula arahan kerja tidak melebihi dr masa tamat arahan kerja");
		f.masa_tamatH.focus();
		return false;
	}
	if (masamula == masatamat)
	{
		alert("Pastikan mula arahan kerja tidak sama dgn masa tamat arahan kerja");
		f.masa_tamatH.focus();
		return false;
	}	
}


//  End -->
</script>

</head>
<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%=color4%>">
<% 
proses1 = request.form("proses1")
if proses1 <> "" then proses="Tambah"
if proses = "" then proses = request.form("proses")
blnthn = request.form("blnthn")
papar
id = request.form("id")
if proses <> "" then	
		if proses="Simpan" then 
			simpan
		elseif proses="Edit" then
			edit(id)
			borang(id)
		elseif proses="Hapus" then 
			rowidH = request.form("rowid")
			hapus(rowidH)
			'response.write("rowid--->"& rowidH)
			borang(id)
			'response.write("id--->"& id)
		elseif proses="Hantar" then 
			'check tahun semasa atau thn semasa + 1
			response.redirect "sp21a.asp?blnthn="&blnthn&""
			c1="select * from payroll.kelepasan_am where to_char(tarikh,'yyyy') like '"& mid(blnthn,3,4) &"'"
			set rc1=objConn.Execute(c1)
			if rc1.bof and rc1.eof then %>
			<script language="JavaScript">
			alert("Maklumat cuti tahunan <%=mid(blnthn,3,4)%> belum ada.  Pendaftaran arahan kerja tidak dapat dibuat\nSila maklumkan pada unit gaji");
			document.location.href = "sp21.asp";
			</script>
			<%
			end if
			borang(id)
		else
			tambah
			borang(id)
		end if
end if

sub tambah()

c1="select to_char(tkh_sah,'dd/mm/yyyy') as tkh_sah from payroll.arahan_kerja where id_ak='"& id &"'"
set cc1 = objConn.Execute(c1)

if not cc1.bof and not cc1.eof then
	if isnull(cc1("tkh_sah")) then
		masa_mulaH = request.form("masa_mulaH")
		masa_mulaM = request.form("masa_mulaM")
		masa_tamatH = request.form("masa_tamatH")
		masa_tamatM = request.form("masa_tamatM")
		tkh = request.form("tkh")
		masa_mula=tkh &" "& masa_mulaH&":"&masa_mulaM
		masa_tamat=tkh &" "& masa_tamatH&":"&masa_tamatM
		masa_mulaKira=masa_mulaH&"."&masa_mulaM
		masa_tamatKira=masa_tamatH&"."&masa_tamatM


		if cdbl(masa_tamatKira)=0 then
				masa_mulatemporary=tkh &" 00:00"
				tkhtemporary=DateAdd("d",1,Cdate(masa_mulatemporary))
				masa_tamat = tkhtemporary
				masa_tamat_i = masa_mulatemporary 
				masa_tamatKira = 24
			else
				masa_tamat_i=masa_tamat
		end if

		'1998-DEC-25 17:30','YYYY-MON-DD HH24:MI'
		'check status cuti
		'1.check public holiday
		'2. check weekend
		r="select tarikh from kelepasan_am where to_char(tarikh,'mm/dd/yyyy') like '"& tkh &"'"
		
		set rs1=objConn.Execute(r)
		'response.write r
		if not rs1.bof and not rs1.eof then
			statusHari="CU"
		else
			'q="select tarikh from sabtu where to_char(tarikh,'mm/dd/yyyy') like '"& tkh &"'"
			q="select decode(to_char(to_date('"& tkh &"','mm/dd/yyyy'),'DY'),'SUN','SA','SAT','SA','P') as statusHari from dual "
			set rs=objConn.Execute(q)
			if not rs.bof and not rs.eof then
				statusHari=rs("statusHari")
			end if
		end if


		jam_M=0
		minit_M=0
		jam_S=0
		minit_S=0
		'get timeframe
		if Cdbl(masa_mulaKira)>=6 and Cdbl(masa_tamatKira) <=22 then
			'cth 6.00 -> 22.00
			'status1="S"
			'status2="S"
			
			'cari beza masa
			jum_minit = datediff("n", Cdate(masa_mula), Cdate(masa_tamat))
			jam_S = Fix(jum_minit / 60)
			minit_S = (jum_minit mod 60)/60
			
			if statusHari="P" then	
				kadar_S=1.125
				kadar_M=0
			elseif statusHari="SA" then
				kadar_S=1.25
				kadar_M=0
			elseif statusHari="CU" then
				kadar_S=1.75
				kadar_M=0
			end if
			'call sub utk insert record
			call insertTkh(id,"S",statusHari,jam_S,minit_S,jam_M,minit_M,kadar_S,kadar_M,tkh,masa_mula,masa_tamat_i,kumpulan)
			
		elseif Cdbl(masa_mulaKira)<6 and Cdbl(masa_tamatKira) > 22 then
			'cth=5.00 -> 23.00
			'status1="M"
			'status2="S"
			'status3="M"
			
			'cari beza masa 
			jum_minit_status1 = datediff("n", Cdate(masa_mula), Cdate(tkh &" 06:00"))
			jum_minit_status3 = datediff("n", Cdate(tkh &" 22:00"), Cdate(masa_tamat))
			jum_minit = Cint(jum_minit_status1) + Cint(jum_minit_status3)
			'response.write("<br>--->jum minit malam"& jum_minit)
			jam_M = Fix(jum_minit / 60)
			minit_M = (jum_minit mod 60)/60
			jum_minit = datediff("n", Cdate(tkh &" 06:00"), Cdate(tkh &" 22:00"))
			'response.write("<br>--->jum minit siang"& jum_minit)
			jam_S = Fix(jum_minit / 60)
			minit_S = (jum_minit mod 60)/60
			
			if statusHari="P" then
				kadar_M=1.25
				kadar_S=1.125		
			elseif statusHari="SA" then
				kadar_S=1.25
				kadar_M=1.50
			elseif statusHari="CU" then
				kadar_S=1.75
				kadar_M=2.00
			end if
		
			'call sub utk insert record
			call insertTkh(id,"A",statusHari,jam_S,minit_S,jam_M,minit_M,kadar_S,kadar_M,tkh,masa_mula,masa_tamat_i,kumpulan)
		
			
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
			
			if statusHari="P" then
				kadar_M=1.25
				kadar_S=0		
			elseif statusHari="SA" then
					kadar_S=0
					kadar_M=1.50
			elseif statusHari="CU" then
					kadar_S=0
					kadar_M=2.00
			end if
			'call sub utk insert record
			call insertTkh(id,"M",statusHari,jam_S,minit_S,jam_M,minit_M,kadar_S,kadar_M,tkh,masa_mula,masa_tamat_i,kumpulan)
			
		elseif Cdbl(masa_mulaKira)<6 and Cdbl(masa_tamatKira) < 6 then
			'cth 01.00 -> 4.00
			'status1="M"
			'status2="M"
			
			'cari beza masa
			jum_minit = datediff("n", Cdate(masa_mula), Cdate(masa_tamat))
			jam_M = Fix(jum_minit / 60)
			minit_M = (jum_minit mod 60)/60	
			
			if statusHari="P" then
				kadar_M=1.25
				kadar_S=0		
			elseif statusHari="SA" then
					kadar_S=0
					kadar_M=1.50
			elseif statusHari="CU" then
					kadar_S=0
					kadar_M=2.00
			end if
			'call sub utk insert record
			call insertTkh(id,"M",statusHari,jam_S,minit_S,jam_M,minit_M,kadar_S,kadar_M,tkh,masa_mula,masa_tamat_i,kumpulan)
			
		elseif Cdbl(masa_mulaKira)>=6 and Cdbl(masa_tamatKira) > 22 then
			'cth 017.00 -> 23.00
			'status1="S"
			'status2="M"
			
			'cari beza masa
			jum_minit = datediff("n", Cdate(masa_mula), Cdate(tkh &" 22:00"))
			jam_S = Fix(jum_minit / 60)
			minit_S = (jum_minit mod 60)/60
			
			jum_minit = datediff("n", Cdate(tkh &" 22:00"), Cdate(masa_tamat)) 'Asma
			jam_M = Fix(jum_minit / 60)
			minit_M = (jum_minit mod 60)/60	
			
			if statusHari="P" then
				kadar_M=1.25
				kadar_S=1.125		
			elseif statusHari="SA" then
					kadar_S=1.25
					kadar_M=1.50
			elseif statusHari="CU" then
					kadar_S=1.75
					kadar_M=2.00
			end if
			'call sub utk insert record
			call insertTkh(id,"A",statusHari,jam_S,minit_S,jam_M,minit_M,kadar_S,kadar_M,tkh,masa_mula,masa_tamat_i,kumpulan)
			
		elseif Cdbl(masa_mulaKira)<6 and Cdbl(masa_tamatKira) <= 22 then
			'cth 05.00 -> 18.00
			'status1="M"
			'status2="S"
			
			'cari beza masa
			jum_minit = datediff("n", Cdate(masa_mula), Cdate(tkh &" 06:00"))
			jam_M = Fix(jum_minit / 60)
			minit_M = (jum_minit mod 60)/60	
			
			jum_minit = datediff("n", Cdate(tkh &" 06:00"), Cdate(masa_tamat))
			jam_S = Fix(jum_minit / 60)
			minit_S = (jum_minit mod 60)/60
			
			if statusHari="P" then
				kadar_M=1.25
				kadar_S=1.125		
			elseif statusHari="SA" then
				kadar_S=1.25
				kadar_M=1.50
			elseif statusHari="CU" then
				kadar_S=1.75
				kadar_M=2.00
			end if
			
			'call sub utk insert record
			call insertTkh(id,"A",statusHari,jam_S,minit_S,jam_M,minit_M,kadar_S,kadar_M,tkh,masa_mula,masa_tamat_i,kumpulan)
		else
			response.write("Proses tidak lengkap bagi masa "& masa_mula &" hingga "& masa_tamat &"<br>Sila hubungi system administrator<br><br>")
		
		end if
	else %>
		<script language="javascript">
        alert("Proses tidak dpt diteruskan kerana arahan kerja <%=id%> sudah disahkan pada <%=cc1("tkh_sah")%>.");
        </script>
	<%
    end if
end if
end sub

sub insertTkh(id,statusMasa,statusHari,jam_S,minit_S,jam_M,minit_M,kadar_S,kadar_M,tkh,masa_mula,masa_tamat,kumpulan)
if cdbl(jam_S) < 0 or cdbl(minit_S) < 0 or Cdbl(jam_M) < 0 or cdbl(minit_M) < 0 then
%>
<script language="javascript">
alert("Waktu tidak sah.  Sila isi semula");
</script>
<%
else 
blnthn_sub = mid(tkh,1,2) & mid(tkh,7,4)

'delete utk elak duplicate record
dd="delete from payroll.jadual_ot where id_ak='"& id &"' and masa_mula=to_date('"& masa_mula &"','mm/dd/yyyy HH24:MI') "
dd = dd & "and masa_tamat=to_date('"& masa_tamat &"','mm/dd/yyyy HH24:MI')"
objConn.Execute(dd)

'insert record 
i1="insert into payroll.jadual_ot (id_ak,tkh_ot,jam_S,minit_S,jam_M,minit_M,masa_mula,masa_tamat,kategori_masa,"
i1 = i1 & "kategori_hari,kumpulan,kadar_S,kadar_M) values "
i1 = i1 & "('"& id &"',to_date('"& tkh &"','mm/dd/yyyy'),"& jam_S &",round("& minit_S &",2),"& jam_M &",round("& minit_M &",2),"
i1 = i1 & "to_date('"& masa_mula &"','mm/dd/yyyy HH24:MI'),"
i1 = i1 & "to_date('"& masa_tamat &"','mm/dd/yyyy HH24:MI'),'"& statusMasa&"','"& statusHari &"',null,"& kadar_S &","& kadar_M &")"
'response.write(i1)
objConn.Execute(i1)

'jamtotal_S=Cint(jam_S) + Cdbl(minit_S)

'jamtotal_M=Cint(jam_M) + Cdbl(minit_M)
end if
end sub

sub edit(id)

c1="select to_char(tkh_sah,'dd/mm/yyyy') as tkh_sah from payroll.arahan_kerja where id_ak='"& id &"'"
set cc1 = objConn.Execute(c1)

if not cc1.bof and not cc1.eof then
	if isnull(cc1("tkh_sah")) then	
	
		penyelia = request.form("penyelia")
		ptj = request.form("ptj")
		unit = request.form("unit")
		keterangan = replace(request.form("keterangan"),"'","''")
		nota = replace(request.form("nota"),"'","''")	
	 
		 g="select to_char(sysdate,'ddmmyyyy')||' 00:00' as todaydate from dual"
		 set og = objConn.Execute(g)
		 todaydate = og("todaydate")
		 
		 'response.write("<b>todaydate--->"& todaydate)
		 
		 i1="update payroll.arahan_kerja set penyelia="& penyelia &",unit='"& unit &"',ptj='"& ptj &"',"
		 i1= i1 & "keterangan='"& keterangan &"',nota='"& nota &"',tkh_input=to_date('"& todaydate &"','ddmmyyyy HH24:MI') "
		 i1 = i1 & "where id_ak='"& id &"'"
		 objConn.Execute(i1)
	else %>
		<script language="javascript">
        alert("Proses tidak dpt diteruskan kerana arahan kerja <%=id%> sudah disahkan pada <%=cc1("tkh_sah")%>.");
        </script>
	<%
    end if
end if
end sub

sub simpan()
bln = request.form("bln")
thn = request.form("thn")
penyelia = request.form("penyelia")
ptj = request.form("ptj")
no_pekerja = request.form("no_pekerja")
unit = request.form("unit")
keterangan = replace(request.form("keterangan"),"'","''")
nota = replace(request.form("nota"),"'","''")

q0="select * from payroll.arahan_kerja where bulan="& bln &" and tahun="& thn &" and penyelia="& penyelia &" and "
q0 = q0 & "ptj="& ptj &" and pemohon="& no_pekerja &" and unit like '"& unit &"' and upper(keterangan) like upper('"& keterangan &"') and "
q0 = q0 & "upper(nota) like upper('"& nota &"') and to_char(tkh_input,'ddmmyyyy') like to_char(sysdate,'ddmmyyyy')"
set oq0 = objConn.Execute(q0)

if oq0.bof and oq0.eof then  'utk elak problem refresh
q1="select max(substr(id_ak,10,4)) siri from payroll.arahan_kerja "
q1 = q1 & " where id_ak like  '"&ptj&"'||'"&thn&"'||'"&bln&"'||'%' "
set oq1 = objConn.Execute(q1)

	if oq1.bof and oq1.eof then	 
        d1 = " select '"& ptj &"'||'"& thn &"'||'"& bln &"'||'0001' dok from dual "
        Set od1 = objConn.Execute(d1)	
        
        id = od1("dok")

     else
        qsiri = oq1("siri")
        if qsiri <> "" then
           zsiri = cdbl(qsiri) + 1
        else
           zsiri = "1"
        end if      

        d1 = " select '"& ptj &"'||'"& thn &"'||'"& bln &"'||lpad('"&zsiri&"',4,'0') dok from dual "
        Set od1 = objConn.Execute(d1)	
        
        id = od1("dok")
     end if  
	 
	 g="select to_char(sysdate,'ddmmyyyy')||' 00:00' as todaydate from dual"
	 set og = objConn.Execute(g)
	 todaydate = og("todaydate")
	 
	 'response.write("<b>todaydate--->"& todaydate)
	 
	 i1="insert into payroll.arahan_kerja (id_ak,penyelia,pemohon,pengesahan,tkh_sah,unit,ptj,keterangan,nota,tkh_input,tahun,bulan) values "
	 i1 = i1 & "('"& id &"',"& penyelia &","&  no_pekerja &",'T',null,'"& unit &"','"& ptj &"','"& keterangan &"',"
	 i1 = i1 & "'"& nota &"',to_date('"& todaydate &"','ddmmyyyy HH24:MI'),"& thn &","& bln &")"
	 objConn.Execute(i1)
	 borang(id)
%>
<script language="JavaScript">
	alert("Selesai simpan maklumat arahan kerja.\nTeruskan proses dengan mandaftar jadual waktu arahan kerja");
	document.masa0.tkh.focus();
</script>
<%	 
end if     
end sub 

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
<form name="test" method="post" action="sp21.asp" onSubmit="return check1(this)">
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
<tr><td width="15%"><b>Bulan/Tahun(mmyyyy):</b> </td><td><input type="text" name="blnthn" size="5" maxlength="6" value="<%=blnthn%>"></td></tr>
<tr align="center"><td colspan="2"><input type="submit" name="proses" value="Hantar"></td></tr>
</table>
</form>
<%
end sub %>
<%
sub borang(id)
if id <> "" then 
	prosesValue="Edit"
Else
	prosesValue="Simpan"
End if
bln = mid(blnthn,1,2)
thn = mid(blnthn,3,4)

ptj=plokasi
no_pekerja=request.cookies("gnop")
q1="select no_pekerja,upper(nama) as nama,kod_ptj from payroll.otpenyelia where kod_ptj='"& ptj &"' order by no_pekerja asc"
set oq1 = objConn.Execute(q1)

q2="select upper(nama) as nama from payroll.paymas where no_pekerja="& no_pekerja &""
set oq2 = objConn.Execute(q2)
If not oq2.bof and not oq2.eof then
	nama = oq2("nama")
end if

tkh=bln &"/01/"& thn
tkhawalbln=DateAdd("m",1,Cdate(tkh))
tkhlast=DateAdd("d",-1,Cdate(tkhawalbln))
position1 = instr(tkhlast,"/")
harilast = mid(tkhlast,position1 + 1,2)

q3="select penyelia,pemohon,pengesahan,unit,ptj,keterangan,nota from payroll.arahan_kerja where id_ak ='"& id &"'"
set oq3 = objConn.Execute(q3)
if not oq3.eof then
	penyelia = oq3("penyelia")
	ptj = oq3("ptj")
	no_pekerja = oq3("pemohon")
	unit = oq3("unit")
	keterangan = oq3("keterangan")
	nota = oq3("nota")
	pengesahan = oq3("pengesahan")
end if

q4="select count(*) as kira from payroll.pekerja_ot where id_ak='"& id &"'"
set oq4 = objConn.Execute(q4)
if Cdbl(oq4("kira")) > 0 then pekerjaOT="ada"

'q5="select kodlokasi_caw,nama_lokasi from kehadiran.lokasi_caw order by nama_lokasi"
'set oq5 = objConn.Execute(q5)
%>
<br>
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
  <TBODY>
  <form name="test0" method="post" action="sp21.asp" onSubmit="return checkForm1(this)">
    <tr> 
      <td class="hd">No Rujukan Arahan Kerja</td>
      <td><input type="hidden" value="<%=id%>" name="id"><%=id%></td>
    </tr>
    <tr>
      <td class="hd">Bln&nbsp;/Thn</td>
      <td><input type="hidden" name="bln" size="2" maxlength="2" value="<%=bln%>"><%=bln%>
        &nbsp;/&nbsp;
        <input type="hidden" value="<%=thn%>" name="thn" size="4" maxlength="4"><%=thn%></td>
    </tr>
    <tr> 
      <td class="hd">Penyelia</td>
      <td> <select name="penyelia">
          <%
	do while not oq1.eof
	%>
          <option <%if Cdbl(penyelia)=Cdbl(oq1("no_pekerja")) then%>selected <%end if%>value="<%=oq1("no_pekerja")%>"><%=oq1("no_pekerja")%>&nbsp;-&nbsp;<%=oq1("nama")%></option>
          <%oq1.movenext
		loop%>
        </select> </td>
    </tr>
    <tr>
      <td class="hd">Pemohon</td>
      <td><%=nama%> (<%=no_pekerja%>) <input type="hidden" name="no_pekerja" value="<%=no_pekerja%>"></td>
    </tr>
    <tr>
      <td class="hd">Jabatan / Unit</td>
      <td><input type="hidden" name="ptj" value="<%=ptj%>">
        <%=ptj%>&nbsp;/&nbsp;
        <input type="text" name="unit" size="25" maxlength="60" value="<%=unit%>"></td>
    </tr>
    
    <tr> 
      <td class="hd">Keterangan Kerja</td>
      <td><textarea name="keterangan" cols="40" rows="3"/><%=keterangan%></textarea></td>
    </tr>
    <tr>
      <td class="hd"> Nota Tambahan</td>
      <td><textarea name="nota" cols="40" rows="3"/><%=nota%></textarea></td>
    </tr>
    <tr align="center">
      <td colspan="2"><input type="hidden" name="blnthn" value="<%=blnthn%>">
	  <%if pengesahan <> "Y" then%><input type="submit" name="proses" value="<%=prosesValue%>"><%end if%></td>
    </tr>
  </form></TBODY>
</TABLE>
<br>
<%'if kumpulan <> "" then%>
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
  <TBODY> 
  <tr class="hd"> 
    <td colspan="5" >Senarai  Tarikh dan Masa Arahan Kerja</td>
  </tr> 
  <tr style="BACKGROUND-COLOR: #ffffff"><td colspan="5">&nbsp;</td></tr>
  <tr class="hd"> 
    <td align="center">Bil</td>
    <td align="center">Tarikh</td>
    <td align="center">Masa Mula (HH24:MM) </td>
    <td align="center">Masa Tamat (HH24:MM) </td>
    <td align="center">Proses</td>
  </tr>
<form name="masa0" method="post" action="sp21.asp" onSubmit="return checkTambah(this);">
<input type="hidden" name="id" value="<%=id%>">
<input type="hidden" name="blnthn" value="<%=blnthn%>">
  <tr bgcolor="#EFF5F5">
    <td>&nbsp;</td>
    <td><select name="tkh"><option value="">Sila Pilih</option>
    <%
	
	For i = 1 To Cdbl(harilast) 
	if len(i)=1 then 
		display= bln &"/"& "0"& Cstr(i) &"/"& thn
		display1 = "0"& Cstr(i) & "/"& bln &"/"& thn
	else
		display=bln &"/"& Cstr(i) &"/"& thn
		display1 = Cstr(i) & "/"& bln &"/"& thn
	end if
	%>
    <option value="<%=display%>"><%=display1%></option>
        <%Next%>
      </select>
    </td>
    <td ><select name="masa_mulaH"><option value="">Sila Pilih</option>
        <%For i = 0 To 23 
	if len(i)=1 then 
		jam="0"& Cstr(i)
	else
		jam=Cstr(i)
	end if
	%>
        <option value="<%=jam%>"><%=jam%></option>
        <%Next%>
      </select>
      &nbsp;:&nbsp;
      <select name="masa_mulaM"><option value="">Sila Pilih</option>
        <%'For i = 0 To 59 
	'if len(i)=1 then 
		'minit="0"& Cstr(i)
	'else
		'minit=Cstr(i)
	'end if
	'<option value="<%=minit>"><%=minit></option>
	'Next%>
     <option value="00">00</option>
     <option value="30">30</option>
      </select></td>
    <td><select name="masa_tamatH"><option value="">Sila Pilih</option>
      <%For i = 0 To 23 
	if len(i)=1 then 
		jam="0"& Cstr(i)
	else
		jam=Cstr(i)
	end if
	%>
      <option value="<%=jam%>"><%=jam%></option>
      <%Next%>
      </select>
      &nbsp;:&nbsp;
      <select name="masa_tamatM"><option value="">Sila Pilih</option>
        <%'For i = 0 To 59 
	'if len(i)=1 then 
		'minit="0"& Cstr(i)
	'else
		'minit=Cstr(i)
	'end if
	'<option value="<%=minit>"><%=minit></option>
	'Next%>
     <option value="00">00</option>
     <option value="30">30</option>
      </select></td>
    <td><input type="submit" value="Simpan" name="proses1"></td>
    </tr>
    </form>
    <% 
bil = 0
tot_ot = 0 
q7 ="select rowid,to_char(tkh_ot,'dd/mm/yyyy') as tkh_ot, to_char(masa_mula,'HH24:MI') as masa_mula, to_char(masa_tamat,'HH24:MI') as masa_tamat,"
q7 = q7 & "kumpulan,masa_mula as susun , masa_tamat as susun_tamat from payroll.jadual_ot where id_ak='"& id &"' order by susun , susun_tamat"	
set oq7 = ObjConn.Execute(q7)
do while not oq7.eof
bil = bil + 1
bil_pekerja=""
angg_totD=""
kumpulanD=""
masa_mula=""
masa_tamat=""
q8="select count(distinct no_pekerja) as bil_pekerja, nvl(sum(angg_ot_S),0) angg_tot_S, nvl(sum(angg_ot_M),0) angg_tot_M "
q8 = q8 & "from payroll.pekerja_ot where id_ak='"& id &"' and to_char(tkh_ot,'dd/mm/yyyy') like '"& oq7("tkh_ot") &"' "
q8 = q8 & "and to_char(masa_tamat_angg,'HH24:MI') like '"& oq7("masa_tamat") &"' "
q8 = q8 & "and to_char(masa_mula_angg,'HH24:MI') like '"& oq7("masa_mula") &"'"
'response.write(q8)
set oq8 = objConn.Execute(q8)
if not oq8.bof and not oq8.eof then
	bil_pekerja=oq8("bil_pekerja")
	angg_totD = Cdbl(oq8("angg_tot_S")) + Cdbl(oq8("angg_tot_M"))
	tot_ot= cdbl(tot_ot) + Cdbl(oq8("angg_tot_S")) + Cdbl(oq8("angg_tot_M"))
end if

q5="select nama_lokasi from kehadiran.lokasi_caw where kodlokasi_caw like '" & oq7("kumpulan") &"'"
set oq5 = objekConn.Execute(q5)
if not oq5.bof and not oq5.eof then
	kumpulanD=oq5("nama_lokasi")
end if
	%>
  <tr bgcolor="#EFF5F5"> 
  	<td><%=bil%></td>
    <td><%=oq7("tkh_ot")%></td>
    <td><%=oq7("masa_mula")%></td>
    <td><%=oq7("masa_tamat")%></td>     
	<form name="masa<%=bil%>"  method="post" action="sp21.asp"> 
	  <input type="hidden" name="rowid" value="<%=oq7("rowid")%>">
	  <input type="hidden" name="id" value="<%=id%>">
	  <input type="hidden" name="blnthn" value="<%=blnthn%>"> 
	  <input type="hidden" name="masa_mulaD" value="<%=oq7("masa_mula")%>">
	  <input type="hidden" name="masa_tamatD" value="<%=oq7("masa_tamat")%>">
	  <input type="hidden" name="tkh_otD" value="<%=oq7("tkh_ot")%>">  
	  <td>
	    <input type="submit" name="proses" value="Hapus" onClick="return confirm('Anda Pasti Hapus Rekod?')">
      </td>
    </form>
  </tr>  
  <%oq7.movenext
  loop %>
    </tbody> 
</table>
<br>
<% end sub 

sub hapus(rowid)

masa_mula_angg=request.form("masa_mulaD")
masa_tamat_angg=request.form("masa_tamatD")
tkh_ot=request.form("tkh_otD")

c1="select to_char(tkh_sah,'dd/mm/yyyy') as tkh_sah from payroll.arahan_kerja where id_ak='"& id &"'"
set cc1 = objConn.Execute(c1)

if not cc1.bof and not cc1.eof then
	if isnull(cc1("tkh_sah")) then
		d1="delete from payroll.pekerja_ot where id_ak='"& id &"' "
		d1 = d1 & "and to_char(tkh_ot,'dd/mm/yyyy') like '"& tkh_ot &"' "
		d1 = d1 & "and to_char(masa_tamat_angg,'HH24:MI') like '"& masa_tamat_angg &"' "
		d1 = d1 & "and to_char(masa_mula_angg,'HH24:MI') like '"& masa_mula_angg &"'"
		objConn.Execute(d1)
		
		d1="delete from payroll.jadual_ot where rowid like '"& rowid &"'"
		objConn.Execute(d1)
		
		d1="delete from payroll.proses_ot where id_ak like '"& id &"'"
		objConn.Execute(d1)
	else %>
		<script language="javascript">
        alert("Proses tidak dpt diteruskan kerana arahan kerja <%=id%> sudah disahkan pada <%=cc1("tkh_sah")%>.");
        </script>
	<%
    end if
end if
end sub
%>
</body>
</html>