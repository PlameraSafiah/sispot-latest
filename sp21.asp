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


function submitForm1(x)
{
	document.getElementById("funit1").value = x.value
	document.forms["test0"].submit();
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

session("blnthn1") = blnthn


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
			c1="select * from payroll.kelepasan_am where to_char(tarikh,'yyyy') like '"& mid(blnthn,3,4) &"'"
			set rc1=objConn.Execute(c1)
			if rc1.bof and rc1.eof then %>
			<script language="JavaScript">
			alert("Maklumat cuti tahunan <%=mid(blnthn,3,4)%> belum ada.  Pendaftaran arahan kerja tidak dapat dibuat\nSila maklumkan pada unit gaji");
			document.location.href = "sp21.asp";
			</script>
			<%
			end if
			
			''borang(id)
			''response.Redirect("sp21a.asp")
			%>
			<script language="JavaScript">
			alert("Masuk <%=blnthn%>");
			document.location.href = "sp21a.asp"
			</script>
            <%
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
</body>
</html>