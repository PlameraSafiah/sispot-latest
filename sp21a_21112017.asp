<!--#include file="connection.asp" -->
<!--'#include file="spmenu.asp"-->
<%response.buffer=true%>
<%response.cookies("amenu") = "sp21a.asp"%>
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

function submitForm1(x)
{
	document.getElementById("funit1").value = x.value
	document.forms["test0"].submit();
}

function submitForm2(x)
{
	document.getElementById("fpenyelia1").value = x.value
	document.forms["test0"].submit();
}

function submitForm3(x)
{
	document.getElementById("fdaerah1").value = x.value
	document.forms["test0"].submit();
}

function submitForm4(x)
{
	document.getElementById("fparlimen1").value = x.value
	document.forms["test0"].submit();
}

function submitForm5(x)
{
	document.getElementById("fdun1").value = x.value
	document.forms["test0"].submit();
}

function submitForm6(x)
{
	document.getElementById("fkategori1").value = x.value
	document.forms["test0"].submit();
}

function submitForm7(x)
{
	document.getElementById("fketerangan1").value = x.value
	document.forms["test0"].submit();
}

</script>
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

''=========================================================================================================================================

proses1 = request.form("proses1")
if proses1 <> "" then proses="Tambah"
if proses = "" then proses = request.form("proses")
	id = request.form("id")
blnthn = request.querystring("blnthn")
'blnthn = request.form("blnthn")
tkh = request.querystring("tkh")
keterangan= request("keterangan")
kategori  = request("kategori")

	fdaerah1 = request.Form("fdaerah1")
	'response.Write(fdaerah1)
	fparlimen1 = request.Form("fparlimen1")
	fdun1 = request.Form("fdun1")
	funit1 = request.Form("funit1")
	fpenyelia1 = request.Form("fpenyelia1")
	fkategori1 = request.Form("fkategori1")
	fketerangan1 = request.Form("fketerangan1")
    
	fdaerah= request("fdaerah")
	fparlimen = request("fparlimen")
	fdun = request("fdun")	
	funit = request.Form("funit")
	fpenyelia = request.Form("fpenyelia")
	fkategori = request.Form("fkategori")
	fketerangan = request.Form("fketerangan")
	blnthn = session("blnthn1")	
	'response.Write(blnthn)
	bln = mid(blnthn,1,2)
	thn = mid(blnthn,3,4)
	''bln = request.form("bln")
	''thn = request.form("thn")
	nota = replace(request.form("nota"),"'","''")
	ptj = request.form("ptj")
	no_pekerja = request.form("no_pekerja")
	tempat = request.Form("tempat")


	'borang(id)
	'id = request.form("id")
	
	''--------------------untuk displaykan data even page load --------------------------
	if id <> "" then id = replace(id,"'","`") end if
	if bln <> "" then bln = replace(bln,"'","`") end if
	if thn <> "" then thn = replace(thn,"'","`") end if
	if nota <> "" then nota = replace(nota,"'","`") end if
	if fpenyelia <> "" then fpenyelia = replace(fpenyelia,"'","`") end if



	

if proses <> "" then	
'response.Write("proses apa ? : ")
'response.Write(proses)
blnthn = request.form("blnthn")
'papar
id = request.form("id")
		if proses="Simpan" then 
			simpan
			borang(id)
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
else 
borang(id)  'saf tambah 12/10/2017 utk papar balik lepas simpan		
		
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
		r="select tarikh from payroll.kelepasan_am where to_char(tarikh,'dd/mm/yyyy') like '"& tkh &"'"
		set rs1=objConn.Execute(r)
		if not rs1.bof and not rs1.eof then
			statusHari="CU"
		else
			'q="select tarikh from sabtu where to_char(tarikh,'mm/dd/yyyy') like '"& tkh &"'"
			q="select decode(to_char(to_date('"& tkh &"','dd/mm/yyyy'),'DY'),'SUN','SA','SAT','SA','P') as statusHari from dual "
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
			status1="M"
			status2="S"
			status3="M"
			
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
dd="delete from payroll.jadual_ot where id_ak='"& id &"' and masa_mula=to_date('"& masa_mula &"','dd/mm/yyyy HH24:MI') "
dd = dd & "and masa_tamat=to_date('"& masa_tamat &"','dd/mm/yyyy HH24:MI')"
objConn.Execute(dd)
session.Abandon '---kill session
'insert record 
i1="insert into payroll.jadual_ot (id_ak,tkh_ot,jam_S,minit_S,jam_M,minit_M,masa_mula,masa_tamat,kategori_masa,"
i1 = i1 & "kategori_hari,kumpulan,kadar_S,kadar_M) values "
i1 = i1 & "('"& id &"',to_date('"& tkh &"','dd/mm/yyyy'),"& jam_S &",round("& minit_S &",2),"& jam_M &",round("& minit_M &",2),"
i1 = i1 & "to_date('"& masa_mula &"','dd/mm/yyyy HH24:MI'),"
i1 = i1 & "to_date('"& masa_tamat &"','dd/mm/yyyy HH24:MI'),'"& statusMasa&"','"& statusHari &"',null,"& kadar_S &","& kadar_M &")"
'response.write(i1)
objConn.Execute(i1)
session.Abandon '---kill session
'jamtotal_S=Cint(jam_S) + Cdbl(minit_S)

'jamtotal_M=Cint(jam_M) + Cdbl(minit_M)
end if
end sub

sub edit(id)
'id=request.cookies("id")
'id=1012017100044
id=request.form("id")
response.write id
bln = request.form("bln")
thn = request.form("thn")


c1="select to_char(tkh_sah,'dd/mm/yyyy') as tkh_sah from payroll.arahan_kerja where id_ak= '"& id &"'"
set cc1 = objConn.Execute(c1)
response.write c1
if not cc1.bof and not cc1.eof then
	if isnull(cc1("tkh_sah")) then	
	
		'penyelia = request.form("penyelia")
'		ptj = request.form("ptj")
'		unit = request.form("unit")
'		kategori = request.form("kategori")
'		keterangan = request.form("keterangan")
'		keterangan = replace(request.form("keterangan"),"'","''")
'		nota = replace(request.form("nota"),"'","''")
'		daerah = request.form("daerah")
'		parlimen = request.form("parlimen")
'		dun = request.form("dun")	
'		tempat = request.Form("tempat")

        penyelia = replace(request.form("penyelia"),"'","''")
		ptj = replace(request.form("ptj"),"'","''")
		unit = replace(request.form("unit"),"'","''")
		kategori = replace(request.form("kategori"),"'","''")
		'k'eterangan = request.form("keterangan")
		keterangan = replace(request.form("keterangan"),"'","''")
		nota = replace(request.form("nota"),"'","''")
		daerah = replace(request.form("daerah "),"'","''")
		parlimen = replace(request.form("parlimen"),"'","''")
		dun = replace(request.form("dun"),"'","''")
		tempat = replace(request.form("tempat"),"'","''")
	 
		 g="select to_char(sysdate,'ddmmyyyy')||' 00:00' as todaydate from dual"
		 set og = objConn.Execute(g)
		 todaydate = og("todaydate")
		 

		 'response.write("<b>todaydate--->"& todaydate)
		 
		 i1="update payroll.arahan_kerja set penyelia='"& fpenyelia1 &"',unit='"& funit1 &"',ptj='"& ptj &"',kategori='"& fkategori1 &"',"
		 i1= i1 & "keterangan='"& fketerangan1 &"',nota='"& nota &"',tkh_input=to_date('"& todaydate &"','ddmmyyyy HH24:MI'), "
		 i1= i1 & "daerah='"& fdaerah1 &"',parlimen='"& fparlimen1 &"',dun='"& fdun1 &"',tempat= '"& tempat &"'"
		 i1 = i1 & "where id_ak='"& id &"'"
		 set rsi1=objConn.Execute(i1)
		 session.Abandon '---kill session
		 response.write i1
	else %>
		<script language="javascript">
        alert("Proses tidak dpt diteruskan kerana arahan kerja <%=id%> sudah disahkan pada <%=cc1("tkh_sah")%>.");
        </script>
	<%
    end if
end if
end sub

sub simpan()
'response.Write("<br>nk simpan ni ----- ")
id=request.form("id")
bln = request.form("bln")
thn = request.form("thn")
fpenyelia = request.form("fpenyelia")
funit = request.form("funit")
fkategori = request.form("fkategori")
fketerangan = replace(request.form("fketerangan"),"'","''")
fdaerah= request.form("fdaerah")
fparlimen = request.form("fparlimen")
fdun= request.form("fdun")

q0="select * from payroll.arahan_kerja where bulan="& bln &" and tahun="& thn &" and penyelia='"& fpenyelia1 &"' and "
q0 = q0 & "ptj="& ptj &" and pemohon="& no_pekerja &" and unit like '"& funit1 &"' and upper(keterangan) like upper('"& fketerangan1 &"') and "
q0 = q0 & "upper(nota) like upper('"& nota &"') and to_char(tkh_input,'ddmmyyyy') like to_char(sysdate,'ddmmyyyy')and"
q0 = q0 & " kategori='"& fkategori1 &"' and daerah='"& fdaerah1 &"' and parlimen='"& fparlimen1 &"' and dun='"& fdun1 &"'  "
'response.Write(q0) & "<br>"
set oq0 = objConn.Execute(q0)


if oq0.eof then  'utk elak problem refresh
q1="select max(substr(id_ak,10,4)) siri from payroll.arahan_kerja "
q1 = q1 & " where id_ak like  '"&ptj&"'||'"&thn&"'||'"&bln&"'||'%' "
'response.Write(q1) & "<br>"
set oq1 = objConn.Execute(q1)

	if  oq1.eof then	 
	
        d1 = " select '"& ptj &"'||'"& thn &"'||'"& bln &"'||'0001' dok from dual "
        Set od1 = objConn.Execute(d1)	
        'response.Write(d1) & "<br>"
        id = od1("dok")
		'response.Write(id)
		'response.Write "<br>"

     else
        qsiri = oq1("siri")
		'response.Write(qsiri)
        if qsiri <> "" then
           zsiri = cdbl(qsiri) + 1
        else
           zsiri = "1"
        end if      

        d2 = " select '"& ptj &"'||'"& thn &"'||'"& bln &"'||lpad('"&zsiri&"',4,'0') dok from dual "
		'response.Write(d2) & "<br>"
        Set od2 = objConn.Execute(d2)	
        
        id = od2("dok")
		'response.Write(id)
		'response.Write "<br>"
     end if  
	 
	 g="select to_char(sysdate,'ddmmyyyy')||' 00:00' as todaydate from dual"
	 set og = objConn.Execute(g)
	 todaydate = og("todaydate")
	 
	 'response.write("<b>todaydate--->"& todaydate)
	 
	 i1="insert into payroll.arahan_kerja (id_ak,penyelia,pemohon,pengesahan,tkh_sah,unit,ptj,kategori,keterangan,nota,tkh_input,tahun,bulan,daerah,parlimen,dun,tempat) values "
	 i1 = i1 & "('"& id &"','"& fpenyelia1 &"','"& no_pekerja &"','T',null,'"& funit1 &"','"& ptj &"','"& fkategori1 &"','"& fketerangan1 &"',"
	 i1 = i1 & "'"& nota &"',to_date('"& todaydate &"','ddmmyyyy HH24:MI'),'"& thn &"','"& bln &"','"& fdaerah1 &"','"& fparlimen1 &"','"& fdun1 &"','"& tempat &"' )"
	 
	 'response.write " insert masuk : " & i1
	 objConn.Execute(i1)
	 
	 session.Abandon '---kill session
	'borang(id)
	 
%>
<script language="JavaScript">
	alert("Selesai simpan maklumat arahan kerja.\nTeruskan proses dengan mandaftar jadual waktu arahan kerja");
	document.masa0.tkh.focus();
</script>
<%	 
end if     
end sub 


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
'q1="select no_pekerja,upper(nama) as nama,kod_ptj from payroll.otpenyelia where kod_ptj='"& ptj &"' order by no_pekerja asc"
a1="select no_pekerja,id_unit from payroll.unit_kakitangan where no_pekerja='"& no_pekerja &"' "
set rsa1 = objConn.Execute(a1)
'response.Write(a1)
if not rsa1.eof then
id_unit=rsa1("id_unit")

end if




q2="select upper(nama) as nama from payroll.paymas where no_pekerja="& no_pekerja &""
set oq2 = objConn.Execute(q2)
If not oq2.bof and not oq2.eof then
	nama = oq2("nama")
end if

'tkh=bln &"/01/"& thn
'tkhawalbln=DateAdd("m",1,Cdate(tkh))
'tkhlast=DateAdd("d",-1,Cdate(tkhawalbln))
'position1 = instr(tkhlast,"/")
'harilast = mid(tkhlast,position1 + 1,2)


tkh= "01/"& bln &"/"& thn  'saf bukak 13/10/2017
'tkh ="01/10/2017"
tkhawalbln=DateAdd("m",1,Cdate(tkh))
'response.write tkhawalbln
tkhlast=DateAdd("d",-1,Cdate(tkhawalbln))
'response.write tkhlast
'position1 = instr(tkhlast,"/")
'harilast = mid(tkhlast,position1 + 1,2)
harilast =mid(Date,1,2) 'safbukak sbb tarikh ot tak keluar smpai semalam from sysdate
'response.write harilast
'response.write harilast
'response.write tkhawalbln

'q3="select id_ak,penyelia,pemohon,pengesahan,unit,ptj,kategori,keterangan,nota,daerah,parlimen,dun,tempat from payroll.arahan_kerja where id_ak ='"& id &"'"
'set oq3 = objConn.Execute(q3)
'response.write q3
'if not oq3.eof then
'    id= oq3("id_ak")
'	fpenyelia = oq3("penyelia")
'	ptj = oq3("ptj")
'	no_pekerja = oq3("pemohon")
'	funit = oq3("unit")
'	fkategori = oq3("kategori")
'	fketerangan = oq3("keterangan")
'	nota = oq3("nota")
'	pengesahan = oq3("pengesahan")
'	fdaerah = oq3("daerah")
'	fparlimen = oq3("parlimen")
'	fdun = oq3("dun")
'	tempat = oq3("tempat")
	
	
'response.Write(q3)

'end if


	'------------funit
	if funit1 <> "" then 
	ffunit = funit1
	session("funit2") = funit1
	end if
	
	if funit <> "" then
		ffunit1 = funit
		session("funit2") = funit
	end if 
	
	if ffunit <> "" then
	
	ffunit1 = cstr(ffunit)
	session("funit2") = ffunit
	end if
	
	'------------fpenyelia
	if fpenyelia1 <> "" then 
	ffpenyelia = fpenyelia1
	session("fpenyelia2") = fpenyelia1
	end if
	
	if fpenyelia <> "" then
		ffpenyelia1 = fpenyelia
		session("fpenyelia2") = fpenyelia
	end if 
	
	if ffpenyelia <> "" then
	
	ffpenyelia1 = cstr(ffpenyelia)
	session("fpenyelia2") = ffpenyelia
	end if
	
	
	
	'---------kategori
	if fkategori1 <> "" then 
		ffkategori = fkategori1
		session("fkategori2") = fkategori1
	end if
	
	if fkategori <> "" then
		ffkategori1 = fkategori
		session("fkategori2") = fkategori
	end if 
	
	if ffkategori <> "" then
		ffkategori1 = cstr(ffkategori)
		session("fkategori2") = ffkategori
	end if
	
	'---------keterangan kerja
	if fketerangan1 <> "" then 
	ffketerangan = fketerangan1
	session("fketerangan2") = fketerangan1
	end if
	
	if fketerangan <> "" then
	
	ffketerangan1 = fketerangan
	session("fketerangan2") = fketerangan
	end if 
	
	if ffketerangan <> "" then
	
	ffketerangan1 = cstr(ffketerangan)
	session("fketerangan2") = ffketerangan
	end if
	
	
	''---------------daerah
	
	if fdaerah1 <> "" then 
	ffdaerah = fdaerah1
	session("fdaerah2") = fdaerah1
	end if
	
	if fdaerah <> "" then
	
	ffdaerah1 = fdaerah
	session("fdaerah2") = fdaerah
	end if 
	
	if ffdaerah <> "" then
	
	ffdaerah1 = cstr(ffdaerah)
	session("fdaerah2") = ffdaerah
	end if
	
	
	''---------------parlimen
	
	if fparlimen1 <> "" then 
	ffparlimen = fparlimen1
	session("fparlimen2") = fparlimen1
	end if
	
	if fparlimen <> "" then
	
	ffparlimen1 = fparlimen
	session("fparlimen2") = fparlimen
	end if 
	
	if ffparlimen <> "" then
	
	ffparlimen1 = cstr(ffparlimen)
	session("fparlimen2") = ffparlimen
	end if
	
	''---------------dun
	
	if fdun1 <> "" then 
	ffdun = fdun1
	session("fdun2") = fdun1
	end if
	
	if fdun <> "" then
	
	ffdun1 = fdun
	session("fdun2") = fdun
	end if 
	
	if ffdun <> "" then
	
	ffdun1 = cstr(ffdun)
	session("fdun2") = ffdun
	end if




q4="select count(*) as kira from payroll.pekerja_ot where id_ak='"& id &"'"
set oq4 = objConn.Execute(q4)
if Cdbl(oq4("kira")) > 0 then pekerjaOT="ada"




q5="select kodlokasi_caw,nama_lokasi from kehadiran.lokasi_caw order by nama_lokasi"
set oq5 = objConn.Execute(q5)
%>
<br>
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
  <TBODY>
  <form name="test0" action="sp21a.asp" method="post" onSubmit="return check(this,blnthn,event)">
    <tr> 
      <td class="hd">No Rujukan Arahan Kerja</td>
      <td><%=id%></td>
    </tr>
    <tr>
      <td class="hd">Bln&nbsp;/Thn</td>
      <td><input type="hidden" name="bln" size="2" maxlength="2" value="<%=bln%>"><%=bln%>
        &nbsp;/&nbsp;
        <input type="hidden" value="<%=thn%>" name="thn" size="4" maxlength="4"><%=thn%></td>
    </tr>
    
    
    <tr>
      <td class="hd">Jabatan / Unit</td>
      <td><input type="hidden" name="ptj" value="<%=ptj%>">
        <%=ptj%>&nbsp;/&nbsp;
       <!--<input type="text" name="unit" size="25" maxlength="60" value="<%'=unit%>"></td>-->
       
     <%   '--------------new

     zv1 = "select id_unit,nama_unit from payroll.unit_sispot where ptj='"& ptj &"'"
     Set rszv1 = objConn.Execute(zv1) 

	%>
           
           <select name="funit" id="funit" onChange="submitForm1(this)" >
           <option value="" size="30">Pilih Unit -- </option>
           
           <%

	
	 do while not rszv1.eof 
		   
		   kod_unit = rszv1("id_unit")
		   nama_unit = rszv1("nama_unit")
		   
		   sSel = ""
		   
		  if kod_unit = ffunit1  then sSel = " selected=""selected"""
		   
		   %>
           
           <option value="<%=kod_unit%>" <%=sSel%>> <%=kod_unit%> - <%=nama_unit%> </option>
           
           <% rszv1.movenext
		   loop
		   %>
           </select>
		   
		  <input type="hidden" name="funit1" value="<%=session("funit2")%>" id="funit1"> 
          </td>
    </tr>
    
    
    <tr> 
      <td class="hd">Penyelia</td>
      <td> 

     
           <select name="fpenyelia" id="fpenyelia" onChange="submitForm2(this)" >
           <option value="" size="30">Pilih Penyelia -- </option>
           
     <%
	 '--------------new
     zv1 = "select a.nama_penyelia,b.no_pekerja,b.nama from payroll.unit_sispot_penyelia a left join payroll.paymas b "
	 zv1 = zv1 & "on a.nama_penyelia = b.no_pekerja where id_unit = '"& ffunit1 &"' and ptj='"& ptj &"' "
     Set rszv1 = objConn.Execute(zv1) 
	 
	 do while not rszv1.eof 
		   
		   kod_penyelia = rszv1("nama_penyelia")
		   nama_p = rszv1("nama")
		   sSel = ""
		   
		  if kod_penyelia = ffpenyelia1  then sSel = " selected=""selected"""
	 %>
           
           <option value="<%=kod_penyelia%>" <%=sSel%>> <%=kod_penyelia%> - <%=nama_p%> </option>
           
           <% rszv1.movenext
		   loop
		   %>
           </select>
		   
		  <input type="hidden" name="fpenyelia1" value="<%=session("fpenyelia2")%>" id="fpenyelia1"> 
          
          
  </td>
    </tr>
    <tr>
      <td class="hd">Pemohon</td>
      <td><%=nama%> (<%=no_pekerja%>) <input type="hidden" name="no_pekerja" value="<%=no_pekerja%>"></td>
    </tr>
    
    <% 
	
	%>
     <TR bgcolor="<%=color1%>" > 
  <TD  align="left" bgcolor="<%=color1%>" width="40%"><b>Kategori</b></TD>
  <td valign="center" bgcolor="F4FAFA">
  <select name="fkategori" id="fkategori" onChange="submitForm6(this)">
           <option value="">Pilih Kategori -- </option>
           
           <% 
		   zv3 = "select distinct(kategori) kod from payroll.keterangan_kerja order by kod"
   		   Set rszv3 = objConn.Execute(zv3) 
	 
		   do while not rszv3.eof 
		   
		   kodkategori = rszv3("kod")

		   sSel = ""
		 
		 	if kodkategori = fkategori1  then sSel = " selected=""selected"""
		   %>
           
           <option value="<%=kodkategori%>" <%=sSel%>> <%=kodkategori%> </option>
           
           <% rszv3.movenext
		   loop
		   %>
           </select> <input type="hidden" name="fkategori1" value="<%=fkategori%>">
	</TD>
    </TR>

<TR bgcolor="<%=color1%>" > 
  <TD  align="left" bgcolor="<%=color1%>" width="40%"><b>Keterangan</b></TD>
  <td bgcolor="F4FAFA"> 
  <select name="fketerangan" id="fketerangan" onChange="submitForm7(this)">
           <option value="">Pilih Keterangan Kerja -- </option>
     <%   '--------------new  Keterangan Kerja
     zv = " select upper(keterangan) ket from payroll.keterangan_kerja where kategori = '"& fkategori1 &"' order by ket  "
     Set rszv = objConn.Execute(zv) 


	 do while not rszv.eof 
		   
		   kodket= rszv("ket")
		   		   
		   sSel = ""
		   
		  if kodket = fketerangan1  then sSel = " selected=""selected"""
		   
		   %>
           
           <option value="<%=kodket%>" <%=sSel%>> <%=kodket%> </option>
           
           <% rszv.movenext
		   loop
		   %>
           </select><input type="hidden" name="fketerangan1" value="<%=fketerangan%>">
           
           Sekiranya tiada dalam pilihan,sila daftar <a href="sp110.asp" %> di sini.</a></td>
</tr>

    
    <tr>
      <td class="hd" >Daerah </td><td>
		<select name="fdaerah" id="fdaerah" onChange="submitForm3(this)">
           <option value="">Pilih Daerah -- </option>
           
           <% 
		   zv3 = "select distinct(daerah) kod from hasil.kawasan order by kod"
   		   Set rszv3 = objConn.Execute(zv3) 
	 
		   do while not rszv3.eof 
		   
		   
		   koddaerah = rszv3("kod")
	 

		   sSel = ""
		   
		  if koddaerah = ffdaerah1  then sSel = " selected=""selected"""
		     if koddaerah = "SPU" then
		  		 namadaerah = "SEBERANG PERAI UTARA"
		   elseif koddaerah = "SPT" then
		  		 namadaerah = "SEBERANG PERAI TENGAH"
		   elseif koddaerah = "SPS" then
		  		 namadaerah = "SEBERANG PERAI SELATAN"
		   end if  
		   %>
           
           <option value="<%=koddaerah%>" <%=sSel%>> <%=koddaerah%> - <%=namadaerah%> </option>
           
           <% rszv3.movenext
		   loop
		   %>
           </select> <input type="hidden" name="fdaerah1" value="<%=session("fdaerah2")%>" id="fdaerah1">
	  </td>
    </tr>
	
    <tr> 
      <td class="hd"> Parlimen </td>
      <td>

		<select name="fparlimen" id="fparlimen" onChange="submitForm4(this)">
           <option value="">Pilih Parlimen -- </option>
           
    <%   '--------------new parlimen
     zv = "select DISTINCT a.kod, a.keterangan,b.parlimen,b.daerah from hasil.parlimen a, hasil.kawasan b"
	 zv = zv & " where b.daerah = '"&ffdaerah1&"' and a.kod = b.parlimen order by kod  "
     Set rszv = objConn.Execute(zv) 
	 
	%>
    
           <% do while not rszv.eof 
		   
		   kodparlimen= rszv("kod")
		   namaparlimen = rszv("keterangan")
		   
		   sSel = ""
		   
		  if kodparlimen = ffparlimen1  then sSel = " selected=""selected"""
		   
		   %>
           
           <option value="<%=kodparlimen%>" <%=sSel%>> <%=kodparlimen%> - <%=namaparlimen%> </option>
           
           <% rszv.movenext
		   loop
		   %>
           </select>  <input type="hidden" name="fparlimen1" value="<%=session("fparlimen2")%>" id="fparlimen1" >
      </td>     
     </tr>
        
        <tr> 
        <td class="hd">Dun </td><td>

 <% 
		''-----------------new dun
		 zv = " select kod,UPPER(dun) ket from hasil.kawasan "
         zv = zv & "where daerah = '"& ffdaerah1 &"' and parlimen = '"& ffparlimen1 &"' order by kod" 
         Set rszv = objConn.Execute(zv)
		 %>
        
        <select name="fdun" id="fdun" onChange="submitForm5(this)">
           <option value="">Pilih Dun -- </option>
           
           <% do while not rszv.eof 
		   
		   koddun= rszv("kod")
		   namadun = rszv("ket")
		   
		   sSel = ""
		   
		  if koddun = ffdun1  then sSel = " selected=""selected"""
		   
		   %>
           
           <option value="<%=koddun%>" <%=sSel%>> <%=koddun%> - <%=namadun%> </option>
           
           <% rszv.movenext
		   loop
		   %>
           </select>   <input type="hidden" name="fdun1" value="<%=session("fdun2")%>" id="fdun1" >

</td>   </tr> 

	<tr>
      <td class="hd"> Tempat </td>
      <td><textarea name="tempat" cols="40" rows="3" /><%=tempat%></textarea></td>
    </tr>
    
    <tr>
      <td class="hd"> Nota Tambahan</td>
      <td><textarea name="nota" cols="40" rows="3" /><%=nota%></textarea></td>
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
<form name="masa0" method="post" action="sp21a.asp" onSubmit="return checkTambah(this);">
<input type="hidden" name="id" value="<%=id%>">
<input type="hidden" name="blnthn" value="<%=blnthn%>">
  <tr bgcolor="#EFF5F5">
    <td>&nbsp;</td>
    <td><select name="tkh"><option value="">Sila Pilih</option>
   <%
   For i = 1 To Cdbl(harilast) 
	if len(i)=1 then 
		display="0" & Cstr(i) & "/" & bln &  "/" & thn
		display1 = "0" & Cstr(i) & "/" & bln & "/" & thn
	else
		display= Cstr(i)&"/"& bln &"/"& thn
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
'	if len(i)=1 then 
'		minit="0"& Cstr(i)
'	else
'		minit=Cstr(i)
'	end if 
'	<option value="<%=minit%><%'=minit%></option>
	<%'Next%>
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
        <%For i = 0 To 59 
	if len(i)=1 then 
		minit="0"& Cstr(i)
	else
		minit=Cstr(i)
	end if %>
	<option value="<%=minit%>"><%=minit%></option>
	<%Next%>
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

q5="select nama_lokasi from kehadiran.lokasi_caw where kodlokasi_caw like '" & oq7("kumpulan") &"' "
set oq5 = objConn.Execute(q5)
'response.write q5
if not oq5.bof and not oq5.eof then
	kumpulanD=oq5("nama_lokasi")
end if
	%>
  <tr bgcolor="#EFF5F5"> 
  	<td><%=bil%></td>
    <td><%=oq7("tkh_ot")%></td>
    <td><%=oq7("masa_mula")%></td>
    <td><%=oq7("masa_tamat")%></td>     
	<form name="masa<%=bil%>"  method="post" action="sp21a.asp"> 
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