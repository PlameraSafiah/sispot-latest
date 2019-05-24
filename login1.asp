<%@ Language=VBScript %>
<html>
<script language="javascript">
 function invalid_pwd(a,nilai){
 if (nilai=="tak blh1")
 	alert("Anda bukan pengguna yang sah. Sila pastikan anda memasukkan kata laluan yang betul. Jika anda belum menukar kata laluan sekurang-kurangnya lapan aksara serta kombinasi huruf, nombor dan aksara khas dan anda menggunakan kata laluan dengan huruf kecil, sila masukkan kata laluan dengan huruf besar. Jika anda masih tidak berjaya juga,sila hubungi Bahagian Teknologi Maklumat.");
	if (nilai=="tak blh2")
  alert(a+" Salah. Jika anda belum menukar kata laluan sekurang-kurangnya lapan aksara serta kombinasi huruf, nombor dan aksara khas dan anda menggunakan kata laluan dengan huruf kecil, sila masukkan kata laluan dengan huruf besar. Jika anda masih tidak berjaya juga,sila hubungi Bahagian Teknologi Maklumat untuk mendapatkan katalaluan lama anda. ");
 document.location.href="sistemnet.asp";
   return(true);
 }
</script>
<head>
<title>Sistem Pengurusan Capaian</title>
<link type="text/css" href="menu1.css" REL="stylesheet">
<script type="text/javascript" src="menu1.js"></script>
</head>
<body bgcolor="#FFFFFF" vlink="#0000FF" alink="#0000FF" topmargin="0" leftmargin="0">
<%
const adParamInput = 1
  const adOutput = 2
  const adVarChar = 200
  const adInteger = 3
  const adStateOpen = 1
  const adUseClient = 3
  const adOpenStatic = 3
  const adCmdStoredProc = 4
  const adCmdText = 1
  
  Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"
  
  Set cmdU = Server.CreateObject ("ADODB.Command")
    Set cmdU.ActiveConnection = objConn
	
	Set cmd = Server.CreateObject ("ADODB.Command")
    Set cmd.ActiveConnection = objConn
	Set cmd1 = Server.CreateObject ("ADODB.Command")
    Set cmd1.ActiveConnection = objConn
	
  mm = "select initcap(nama) nama,to_char(sysdate,'dd / mm / yyyy') tkh from majlis.syarikat "
  set rsmm = objconn.execute(mm)
 
pekerja = request.form("nopekerja")
pwd = request.form("password2")
Session ("pass") = pwd
Session("idstaff")= request("nopekerja")
Session ("mula") = index


if pekerja <> "" and pwd <> "" then

'if isnumeric(pwd) then pwd=Cdbl(pwd)

	q =     "select sulit,jabatan,nvl(to_char(tarikh_kemaskini,'mm-dd-yyyy'),'00-00-0000') tarikh_kemaskini from majlis.pengguna "
	q = q & " where no_pekerja = ? and sulit =? "
	
	cmdU.commandText = q
	cmdU.commandType = adCmdText
	
	cmdU.Parameters.Append cmdU.CreateParameter ("no_pekerja", adInteger, adParamInput, , pekerja)
	cmdU.Parameters.Append cmdU.CreateParameter("sulit", adVarChar, adParamInput,Len(pwd) , pwd) 
	set rsq = cmdU.execute
		
	if rsq.eof then
		masuk="tak blh1"
		
	else
	passsulit = rsq("sulit")
	if pwd <> passsulit then
			masuk="tak blh2"
		else
			if isnumeric(pwd) then
				if Cdbl(pekerja) = Cdbl(pwd) then
					session.abandon
		%>
					<script>
					alert("Anda tidak dibenarkan memasuki sistem.\n Sila tukar katalaluan terlebih dahulu");
					document.location.href="sistemnet.asp";
					</script>
		<%
				else
				masuk="blh"
				end if
			else
			
				masuk="blh"
			end if
		end if
	end if	

	  if masuk="blh" then
					'semak tarikh kemaskini samada telah kemaskini atau belum
					tkhKemaskini = rsq("tarikh_kemaskini")
					if tkhKemaskini = "00-00-0000"  then
						session.abandon
		%>
						<script>
						alert("Anda tidak dibenarkan memasuki sistem.\n Sila Kemaskini katalaluan terlebih dahulu");
						document.location.href="sistemnet.asp";
						</script>
		<%
					else
					tkhKemaskininew = cdate(tkhKemaskini)
						bezaHari = DateDiff("d",tkhKemaskininew,Date)
						if bezaHari >90 then
								session.abandon
		%>
						<script>
						alert("Anda tidak dibenarkan memasuki sistem.\nKataLaluan telah cukup 90 hari.\nSila Kemaskini katalaluan terlebih dahulu");
						document.location.href="sistemnet.asp";
						</script>
		<%			
						else
							
					
		 jabatan = cstr(rsq("jabatan"))  
	  	  m = " select no_pekerja,lokasi from payroll.paymas where no_pekerja=?	"
                  m = m & " union "
		  m = m & " select no_pekerja,jabatan lokasi from payroll.tetap_sambilan where no_pekerja=?"
		  m = m & " union "
		  m = m & " select no_pekerja,lokasi from kontrak.paymas where no_pekerja=?"
		  cmd.commandText = m
		  cmd.commandType = adCmdText
	
		  cmd.Parameters.Append cmd.CreateParameter ("no_pekerja", adInteger, adParamInput, , pekerja)
		  cmd.Parameters.Append cmd.CreateParameter ("no_pekerja", adInteger, adParamInput, , pekerja)
		  cmd.Parameters.Append cmd.CreateParameter ("no_pekerja", adInteger, adParamInput, , pekerja)
		  set sm = cmd.execute
				 lokasi = cstr(sm("lokasi"))		 
		if jabatan <> lokasi then		   
		    n = " update majlis.pengguna set jabatan = ? where no_pekerja = ? "
			 cmd1.commandText = n
		  	 cmd1.commandType = adCmdText
			 
			 cmd1.Parameters.Append cmd1.CreateParameter ("jabatan", adInteger, adParamInput, , CInt(lokasi))
		     cmd1.Parameters.Append cmd1.CreateParameter ("no_pekerja", adInteger, adParamInput, , pekerja)
			 set sn = cmd1.execute
		end if	

	     response.cookies("gnop") = pekerja
	     response.cookies("pass") = pwd		 
         response.cookies("pekerjaglobal") = request.form("pekerja")
         response.write "<script language = ""javascript"">"
         response.write "var timeID = setTimeout(""location.href = 'sistem.asp'"",1)"
         response.write "</script>"
         response.end
						end if
					end if
      else
         response.write "<script language = ""javascript"">"
         response.write "var timeID = setTimeout('invalid_pwd(""Password"","""& masuk &""");',1)"
         response.write "</script>"
         response.end
	  end if
  
else
response.redirect"sistemnet.asp"
end if
%>

</body>
</html>